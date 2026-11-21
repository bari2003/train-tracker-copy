/**
 * Train Tracker — Frontend Application
 *
 * Requirements:
 *  1. Show ONLY the selected train on the map along its railway track (all other trains hidden).
 *  2. Train icon is locked to the track coordinate and does not drift when zooming.
 *  3. Hyperactive, smooth, continuous live movement of the train along its physical railway track.
 *  4. Trains NOT present in the database open the dedicated API live status page (matching website style).
 */

'use strict';

// ── Map initialisation ────────────────────────────────────────────────────────
const map = L.map('map', { zoomControl: true }).setView([22.5, 78.9], 5);

// Base tile layers
const osmLayer = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    maxZoom: 19,
    attribution: '© OpenStreetMap contributors'
});

const esriSatellite = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
    maxZoom: 19,
    attribution: 'Tiles &copy; Esri &mdash; Source: Esri, i-cubed, USDA, USGS, AEX, GeoEye, Getmapping, Aerogrid, IGN, IGP, UPR-EGP, and the GIS User Community'
});

// Railway Tracks layer (OpenRailwayMap — physical rails, tracks, signals, stations)
const railwayTracksLayer = L.tileLayer('https://{s}.tiles.openrailwaymap.org/standard/{z}/{x}/{y}.png', {
    maxZoom: 19,
    attribution: 'Map data: © <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors | Style: © <a href="https://www.OpenRailwayMap.org">OpenRailwayMap</a>'
});

// Set default layers
osmLayer.addTo(map);
railwayTracksLayer.addTo(map);

// Add layer control
L.control.layers({
    '🗺️ OpenStreetMap (Default)': osmLayer,
    '🛰️ Satellite View (Esri)': esriSatellite
}, {
    '🚆 Railway Tracks (OpenRailwayMap)': railwayTracksLayer
}, { position: 'topright' }).addTo(map);

// ── State ────────────────────────────────────────────────────────────────────
let stompClient       = null;
let stompSub          = null;
let currentTrainId    = null;
let followMode        = false;

// Single active plotted train elements (all others remain hidden)
let activeRouteLayer  = null;
let activeTrainMarker = null;
let activeStations    = [];

// Hyperactive movement state
let trackPoints       = []; // [[lat, lng], [lat, lng], ...]
let trackDistances    = []; // cumulative meters
let totalTrackMeters  = 0;
let currentDistanceM  = 0;
let animReqId         = null;
let lastFrameTime     = null;
let currentSpeedKmh   = 84; // km/h

// ── DOM refs ─────────────────────────────────────────────────────────────────
const searchInput    = document.getElementById('train-search');
const searchBtn      = document.getElementById('search-btn');
const searchResults  = document.getElementById('search-results');
const infoPanel      = document.getElementById('info-panel');
const followBtn      = document.getElementById('follow-btn');
const stopBtn        = document.getElementById('stop-btn');
const connStatus     = document.getElementById('connection-status');
const dataIndicator  = document.getElementById('data-indicator');

// ── Train Icon (Strict Pixel Anchor, Never Drifts On Zoom) ────────────────────
function createFixedTrainIcon(trainNum, bearing) {
    const rotation = bearing != null ? `transform: rotate(${bearing}deg);` : '';
    return L.divIcon({
        className: 'custom-train-div-icon',
        html: `
            <div class="train-marker-wrapper">
                <div class="train-beacon-pulse"></div>
                <div class="train-marker-icon" id="marker-locomotive" style="${rotation}">🚆</div>
                <div class="train-label-badge">${trainNum}</div>
            </div>
        `,
        iconSize: [36, 36],
        iconAnchor: [18, 18],
        popupAnchor: [0, -22]
    });
}

// ── Bearing calculation between two geographic points ─────────────────────────
function calculateBearing(lat1, lon1, lat2, lon2) {
    const dLon = (lon2 - lon1) * Math.PI / 180;
    const rLat1 = lat1 * Math.PI / 180;
    const rLat2 = lat2 * Math.PI / 180;
    const y = Math.sin(dLon) * Math.cos(rLat2);
    const x = Math.cos(rLat1) * Math.sin(rLat2) - Math.sin(rLat1) * Math.cos(rLat2) * Math.cos(dLon);
    return (Math.atan2(y, x) * 180 / Math.PI + 360) % 360;
}

// Distance in metres between two coordinates (Haversine formula)
function haversineMetres(lat1, lon1, lat2, lon2) {
    const R = 6371000;
    const dLat = (lat2 - lat1) * Math.PI / 180;
    const dLon = (lon2 - lon1) * Math.PI / 180;
    const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
              Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
              Math.sin(dLon / 2) * Math.sin(dLon / 2);
    return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

// ── Hyperactive Track Glider Engine ──────────────────────────────────────────
function initTrackDistances() {
    trackDistances = [0];
    totalTrackMeters = 0;
    for (let i = 0; i < trackPoints.length - 1; i++) {
        const d = haversineMetres(
            trackPoints[i][0], trackPoints[i][1],
            trackPoints[i + 1][0], trackPoints[i + 1][1]
        );
        totalTrackMeters += d;
        trackDistances.push(totalTrackMeters);
    }
}

// Get interpolated coordinate and bearing at a specific distance along the track
function getPositionAtDistance(distM) {
    if (trackPoints.length === 0) return { lat: 20.59, lng: 78.96, bearing: 0 };
    if (trackPoints.length === 1) return { lat: trackPoints[0][0], lng: trackPoints[0][1], bearing: 0 };

    const wrappedDist = distM % totalTrackMeters;

    // Find segment
    let segIdx = 0;
    for (let i = 0; i < trackDistances.length - 1; i++) {
        if (wrappedDist >= trackDistances[i] && wrappedDist <= trackDistances[i + 1]) {
            segIdx = i;
            break;
        }
    }

    const segStartDist = trackDistances[segIdx];
    const segEndDist   = trackDistances[segIdx + 1];
    const segLen       = segEndDist - segStartDist;
    const fraction     = segLen > 0 ? (wrappedDist - segStartDist) / segLen : 0;

    const p1 = trackPoints[segIdx];
    const p2 = trackPoints[segIdx + 1];

    const lat = p1[0] + (p2[0] - p1[0]) * fraction;
    const lng = p1[1] + (p2[1] - p1[1]) * fraction;
    const bearing = calculateBearing(p1[0], p1[1], p2[0], p2[1]);

    return { lat, lng, bearing };
}

// Hyperactive 60 FPS animation frame loop
function startHyperactiveTrackAnimation(trainNumber) {
    if (animReqId) cancelAnimationFrame(animReqId);
    lastFrameTime = performance.now();

    function step(now) {
        if (!activeTrainMarker || trackPoints.length < 2) return;

        const deltaMs = now - lastFrameTime;
        lastFrameTime = now;

        // Speed in m/s with a smooth multiplier for dynamic realism
        const speedMps = (Math.max(currentSpeedKmh, 50) / 3.6);
        currentDistanceM += (speedMps * (deltaMs / 1000) * 8.0); // Hyperactive smooth glider

        const pos = getPositionAtDistance(currentDistanceM);

        activeTrainMarker.setLatLng([pos.lat, pos.lng]);

        // Rotate train icon smoothly to match track curve
        const locomotiveEl = document.getElementById('marker-locomotive');
        if (locomotiveEl) {
            locomotiveEl.style.transform = `rotate(${pos.bearing}deg)`;
        }

        if (followMode) {
            map.panTo([pos.lat, pos.lng], { animate: false });
        }

        animReqId = requestAnimationFrame(step);
    }

    animReqId = requestAnimationFrame(step);
}

function stopHyperactiveAnimation() {
    if (animReqId) {
        cancelAnimationFrame(animReqId);
        animReqId = null;
    }
}

// ── WebSocket ─────────────────────────────────────────────────────────────────
function connectWebSocket(trainId) {
    if (stompClient && stompClient.connected) {
        setConnectionStatus(true);
        subscribeToTrain(trainId);
        return;
    }
    const socket = new SockJS('/ws');
    stompClient = Stomp.over(socket);
    stompClient.debug = null;

    stompClient.connect({}, () => {
        setConnectionStatus(true);
        subscribeToTrain(trainId);
    }, (err) => {
        setConnectionStatus(false);
        console.warn('WebSocket connection retrying in 5s…', err);
        setTimeout(() => {
            if (currentTrainId) connectWebSocket(currentTrainId);
        }, 5000);
    });
}

function subscribeToTrain(trainId) {
    if (stompSub) { stompSub.unsubscribe(); stompSub = null; }
    stompSub = stompClient.subscribe(`/topic/train/${trainId}`, (msg) => {
        try {
            const state = JSON.parse(msg.body);
            updateUI(state);
        } catch (_) {}
    });
}

function disconnectWebSocket() {
    if (stompSub) { stompSub.unsubscribe(); stompSub = null; }
    if (!stompClient || !stompClient.connected) {
        setConnectionStatus(false);
    }
}

// ── Find Closest Meter Distance Along Track For Exact Railway Track Snapping ──
function findClosestDistanceAlongTrack(targetLat, targetLng) {
    if (!trackPoints || trackPoints.length < 2) return 0;
    let minSqDist = Infinity;
    let closestSegIdx = 0;
    let closestT = 0;

    for (let i = 0; i < trackPoints.length - 1; i++) {
        const p1 = trackPoints[i];
        const p2 = trackPoints[i + 1];
        const dx = p2[1] - p1[1];
        const dy = p2[0] - p1[0];
        const segLenSq = dx * dx + dy * dy;
        let t = 0;
        if (segLenSq > 0) {
            t = ((targetLng - p1[1]) * dx + (targetLat - p1[0]) * dy) / segLenSq;
            t = Math.max(0, Math.min(1, t));
        }
        const projLat = p1[0] + t * dy;
        const projLng = p1[1] + t * dx;
        const distSq = (targetLat - projLat) ** 2 + (targetLng - projLng) ** 2;
        if (distSq < minSqDist) {
            minSqDist = distSq;
            closestSegIdx = i;
            closestT = t;
        }
    }

    const segBaseM = trackDistances[closestSegIdx] || 0;
    const nextBaseM = (closestSegIdx + 1 < trackDistances.length) ? trackDistances[closestSegIdx + 1] : totalTrackMeters;
    const segLengthM = nextBaseM - segBaseM;
    return segBaseM + (closestT * segLengthM);
}

// ── UI update from WebSocket message ─────────────────────────────────────────
function updateUI(state) {
    if (!state) return;

    // Info panel
    document.getElementById('info-status').textContent   = state.status || 'RUNNING';
    document.getElementById('info-status').className     = `badge ${state.status || 'RUNNING'}`;

    const locElem = document.getElementById('info-location');
    if (locElem) {
        locElem.textContent = state.currentStation || (state.latitude ? `${state.latitude.toFixed(3)}°, ${state.longitude.toFixed(3)}°` : 'En route');
    }
    const nextHaltElem = document.getElementById('info-next-halt');
    if (nextHaltElem) {
        nextHaltElem.textContent = state.nextHalt || '—';
    }

    const spd = state.speedKmh != null ? state.speedKmh : 84;
    currentSpeedKmh = spd;
    const spdElem = document.getElementById('info-speed');
    if (spdElem) spdElem.textContent = `${Math.round(spd)} km/h`;

    const delayElem = document.getElementById('info-delay');
    if (state.delayMinutes != null) {
        if (state.delayMinutes > 0) {
            delayElem.textContent = `+${state.delayMinutes} min late`;
            delayElem.style.color = '#ff6f6f';
        } else if (state.delayMinutes < 0) {
            delayElem.textContent = `${Math.abs(state.delayMinutes)} min early`;
            delayElem.style.color = '#6fffa0';
        } else {
            delayElem.textContent = 'On time';
            delayElem.style.color = '#6fffa0';
        }
    } else {
        delayElem.textContent = 'On time';
        delayElem.style.color = '#6fffa0';
    }

    document.getElementById('info-updated').textContent  = state.lastUpdated ? formatTime(state.lastUpdated) : 'Just now';
    document.getElementById('info-source').textContent   = state.positionSource || 'LIVE';

    if (state.stale || state.status === 'STALE') {
        dataIndicator.textContent = '⚠ STALE';
        dataIndicator.className   = 'indicator stale';
    } else {
        dataIndicator.textContent = '● LIVE';
        dataIndicator.className   = 'indicator live';
    }

    // Keep active train marker smoothly locked to telemetry position along physical track
    if (activeTrainMarker && totalTrackMeters > 0) {
        let reportedM = null;
        if (state.latitude != null && state.longitude != null) {
            reportedM = findClosestDistanceAlongTrack(state.latitude, state.longitude);
        } else if (state.segmentProgress != null && state.segmentProgress > 0) {
            reportedM = state.segmentProgress * totalTrackMeters;
        }
        if (reportedM != null && Math.abs(currentDistanceM - reportedM) > 2500) {
            currentDistanceM = reportedM;
        }
    }
}

// ── Clear Map (Hides All Previous Trains & Tracks) ───────────────────────────
function clearMap() {
    stopHyperactiveAnimation();

    if (activeTrainMarker) {
        map.removeLayer(activeTrainMarker);
        activeTrainMarker = null;
    }
    if (activeRouteLayer) {
        map.removeLayer(activeRouteLayer);
        activeRouteLayer = null;
    }
    activeStations.forEach(m => map.removeLayer(m));
    activeStations = [];
    trackPoints = [];
    trackDistances = [];
    totalTrackMeters = 0;
    currentDistanceM = 0;
}

// ── Select Train: ONLY Shows This Selected Train, Hides All Others ──────────
async function selectTrain(train) {
    if (currentTrainId && currentTrainId !== train.trainId) {
        fetch(`/api/trains/${currentTrainId}/untrack`, { method: 'POST' }).catch(() => {});
    }

    // Hide everything else on the map!
    clearMap();

    currentTrainId = train.trainId;
    searchResults.innerHTML = '';

    // Populate info panel
    document.getElementById('info-train-name').textContent = train.trainName;
    document.getElementById('info-number').textContent     = train.trainNumber;
    document.getElementById('info-from').textContent       = train.sourceStation?.stationName || train.sourceStation || '—';
    document.getElementById('info-to').textContent         = train.destinationStation?.stationName || train.destinationStation || '—';
    infoPanel.classList.remove('hidden');
    stopBtn.classList.remove('hidden');
    dataIndicator.textContent = '⟳ Connecting…';
    dataIndicator.className   = 'indicator';

    // 1. Fetch & Plot its Dedicated Railway Track + Initial Live Telemetry in Parallel
    try {
        const [routeRes, trackRes] = await Promise.all([
            fetch(`/api/trains/${train.trainId}/route`),
            fetch(`/api/trains/${train.trainId}/track`, { method: 'POST' })
        ]);

        let geoJson = null;
        if (routeRes.ok) {
            geoJson = await routeRes.json();
        }

        let initialState = null;
        if (trackRes.ok) {
            const data = await trackRes.json();
            if (data && data.initialState) {
                initialState = data.initialState;
            }
            connectWebSocket(train.trainId);
        }

        if (geoJson) {
            renderSelectedTrainTrack(geoJson, train.trainNumber, initialState);
        }

        if (initialState) {
            updateUI(initialState);
        }
    } catch (e) {
        console.warn('Error selecting train:', e);
    }
}

// ── Render Selected Train's Route and Place Hyperactive Moving Marker ─────────
function renderSelectedTrainTrack(geoJson, trainNumber, initialState) {
    // Authentic physical railway track layer group (ballast bed casing + glowing active rail path)
    const trackCasing = L.geoJSON(geoJson, {
        style: {
            color: '#0f3460',
            weight: 7,
            opacity: 0.9,
            lineCap: 'round',
            lineJoin: 'round'
        }
    });

    const railCore = L.geoJSON(geoJson, {
        style: {
            color: '#e94560',
            weight: 3.5,
            opacity: 0.95,
            lineCap: 'round',
            lineJoin: 'round'
        },
        pointToLayer: (feature, latlng) => {
            const stationName = feature.properties?.name || feature.properties?.stationName || 'Station';
            const m = L.circleMarker(latlng, {
                radius: 5,
                color: '#ffffff',
                fillColor: '#0f3460',
                fillOpacity: 1,
                weight: 2
            });
            m.bindTooltip(`<b>🚉 ${stationName}</b>`, { direction: 'top' });
            activeStations.push(m);
            return m;
        }
    });

    activeRouteLayer = L.layerGroup([trackCasing, railCore]).addTo(map);

    // Extract coordinate path for hyperactive animation
    trackPoints = [];
    if (geoJson.type === 'LineString' && Array.isArray(geoJson.coordinates)) {
        trackPoints = geoJson.coordinates.map(pt => [pt[1], pt[0]]);
    } else if (geoJson.type === 'FeatureCollection' && Array.isArray(geoJson.features)) {
        const ls = geoJson.features.find(f => f.geometry && f.geometry.type === 'LineString');
        if (ls) {
            trackPoints = ls.geometry.coordinates.map(pt => [pt[1], pt[0]]);
        }
    }

    if (trackPoints.length > 1) {
        initTrackDistances();

        // Determine authentic live starting location directly on physical railway track
        let startDistM = 0;
        if (initialState) {
            if (initialState.latitude != null && initialState.longitude != null) {
                startDistM = findClosestDistanceAlongTrack(initialState.latitude, initialState.longitude);
            } else if (initialState.segmentProgress != null && initialState.segmentProgress > 0) {
                startDistM = Math.min(Math.max(initialState.segmentProgress * totalTrackMeters, 0), totalTrackMeters);
            }
        }

        currentDistanceM = startDistM;
        const initialPos = getPositionAtDistance(currentDistanceM);

        activeTrainMarker = L.marker([initialPos.lat, initialPos.lng], {
            icon: createFixedTrainIcon(trainNumber, initialPos.bearing)
        }).addTo(map);

        activeTrainMarker.bindPopup(`
            <div style="font-family:sans-serif; min-width:160px;">
                <b style="color:#e94560; font-size:1.1rem;">🚆 ${trainNumber}</b><br>
                <b>Live On Railway Track</b><br>
                <small style="color:#6fffa0;">● Real-Time Movement</small>
            </div>
        `);

        // Frame view smoothly: center directly on live train position if running
        if (startDistM > 0) {
            map.setView([initialPos.lat, initialPos.lng], 10, { animate: true });
        } else {
            try {
                map.fitBounds(railCore.getBounds(), { padding: [50, 50] });
            } catch (_) {}
        }

        // Launch hyperactive 60 FPS animation directly from its live running location
        startHyperactiveTrackAnimation(trainNumber);
    }
}

// ── Search & Handling Not-in-Database Trains ───────────────────────────────────
async function searchTrains() {
    const q = searchInput.value.trim();
    if (!q) {
        // Show all DB trains in results list for easy selection
        const res = await fetch('/api/trains');
        const trains = await res.json();
        renderSearchResults(trains, '');
        return;
    }

    const res = await fetch(`/api/trains?q=${encodeURIComponent(q)}`);
    const trains = await res.json();
    renderSearchResults(trains, q);
}

function renderSearchResults(trains, query) {
    searchResults.innerHTML = '';
    const hasMatches = trains && trains.length > 0;

    if (hasMatches) {
        trains.forEach(t => {
            const li = document.createElement('li');
            li.innerHTML = `
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <div>
                        <strong>${t.trainNumber}</strong> — ${t.trainName}
                    </div>
                    <span class="badge RUNNING" style="font-size:0.7rem; background:#1a5c2a;">In Database</span>
                </div>
            `;
            li.addEventListener('click', () => selectTrain(t));
            searchResults.appendChild(li);
        });
    }

    // Check if user's search query is NOT in the database
    const exactMatch = trains.some(t => t.trainNumber.toLowerCase() === query.toLowerCase());

    if (query && !exactMatch) {
        // Not in database -> Show action to open dedicated API live status page
        const apiLi = document.createElement('li');
        apiLi.style.background = '#16213e';
        apiLi.style.borderLeft = '3px solid #e94560';
        apiLi.style.padding = '12px 14px';
        apiLi.innerHTML = `
            <div style="font-weight:700; color:#e94560; margin-bottom:4px;">🌐 Train "${query}" (Not in Database)</div>
            <div style="font-size:0.8rem; color:#94a3b8; margin-bottom:8px;">Open live status & telemetry dashboard via API in new page</div>
            <button style="background:#e94560; color:#ffffff; padding:7px 12px; font-size:0.82rem; border-radius:5px; width:100%; font-weight:700; cursor:pointer; border:none;">
                ↗ Open Live API Status Page
            </button>
        `;
        apiLi.addEventListener('click', () => {
            openExternalTrainPage(query);
        });
        searchResults.appendChild(apiLi);
    }
}

function openExternalTrainPage(trainNumber) {
    const url = `train-live.html?train=${encodeURIComponent(trainNumber)}`;
    window.open(url, '_blank');
}

// ── Follow mode ───────────────────────────────────────────────────────────────
followBtn.addEventListener('click', () => {
    followMode = !followMode;
    followBtn.classList.toggle('active', followMode);
    followBtn.textContent = followMode ? '📍 Following' : 'Follow Train';
});

// ── Stop tracking ─────────────────────────────────────────────────────────────
stopBtn.addEventListener('click', async () => {
    if (!currentTrainId) return;
    await fetch(`/api/trains/${currentTrainId}/untrack`, { method: 'POST' }).catch(() => {});
    disconnectWebSocket();
    clearMap();
    infoPanel.classList.add('hidden');
    stopBtn.classList.add('hidden');
    currentTrainId = null;
    followMode = false;
    followBtn.textContent = 'Follow Train';
    followBtn.classList.remove('active');
    map.setView([22.5, 78.9], 5);
});

// ── Event listeners ───────────────────────────────────────────────────────────
searchBtn.addEventListener('click', () => {
    const q = searchInput.value.trim();
    if (q) {
        fetch(`/api/trains?q=${encodeURIComponent(q)}`)
            .then(r => r.json())
            .then(trains => {
                if (!trains || trains.length === 0) {
                    openExternalTrainPage(q);
                } else if (trains.length === 1) {
                    selectTrain(trains[0]);
                } else {
                    renderSearchResults(trains, q);
                }
            });
    } else {
        searchTrains();
    }
});

searchInput.addEventListener('keydown', e => {
    if (e.key === 'Enter') {
        const q = searchInput.value.trim();
        if (q) {
            fetch(`/api/trains?q=${encodeURIComponent(q)}`)
                .then(r => r.json())
                .then(trains => {
                    if (!trains || trains.length === 0) {
                        openExternalTrainPage(q);
                    } else if (trains.length === 1) {
                        selectTrain(trains[0]);
                    } else {
                        renderSearchResults(trains, q);
                    }
                });
        }
    } else {
        setTimeout(searchTrains, 250);
    }
});

// ── Helpers ───────────────────────────────────────────────────────────────────
function setConnectionStatus(connected) {
    connStatus.textContent = connected ? '🟢 Connected' : '⚪ Standby';
    connStatus.style.color = connected ? '#6fffa0' : '#cbd5e1';
}

function formatTime(isoString) {
    try {
        return new Date(isoString).toLocaleTimeString();
    } catch (_) { return isoString; }
}

// ── Startup ───────────────────────────────────────────────────────────────────
setConnectionStatus(false);
searchTrains();
