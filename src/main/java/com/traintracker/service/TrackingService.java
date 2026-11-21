package com.traintracker.service;

import com.traintracker.config.AppConfig;
import com.traintracker.dto.ProviderTelemetry;
import com.traintracker.dto.TrainStateMessage;
import com.traintracker.model.*;
import com.traintracker.repository.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.messaging.simp.SimpMessagingTemplate;
import org.springframework.scheduling.annotation.Scheduled;
import org.springframework.stereotype.Service;

import java.time.Instant;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * Core tracking service.
 *
 * Responsibilities:
 *  - Periodically poll the external provider for each tracked train.
 *  - Validate telemetry via PositionProcessingService.
 *  - Determine and snap/interpolate the display position.
 *  - Publish processed state to connected WebSocket clients.
 *  - Optionally persist historical snapshots.
 *  - Retain last-known state and expose stale indicator on failure.
 */
@Service
public class TrackingService {

    private static final Logger log = LoggerFactory.getLogger(TrackingService.class);

    private final LiveDataProviderService providerService;
    private final SimulationService simulationService;
    private final PositionProcessingService positionService;
    private final SimpMessagingTemplate messagingTemplate;
    private final TrainRepository trainRepo;
    private final RouteRepository routeRepo;
    private final LiveTrainStateRepository liveStateRepo;
    private final TrackingHistoryRepository historyRepo;
    private final AppConfig config;

    /** Set of train IDs currently being tracked. */
    private final Set<Long> trackedTrains = ConcurrentHashMap.newKeySet();

    /** Update counter per train for sampling history snapshots. */
    private final Map<Long, AtomicInteger> updateCounters = new ConcurrentHashMap<>();

    public TrackingService(LiveDataProviderService providerService,
                           SimulationService simulationService,
                           PositionProcessingService positionService,
                           SimpMessagingTemplate messagingTemplate,
                           TrainRepository trainRepo,
                           RouteRepository routeRepo,
                           LiveTrainStateRepository liveStateRepo,
                           TrackingHistoryRepository historyRepo,
                           AppConfig config) {
        this.providerService = providerService;
        this.simulationService = simulationService;
        this.positionService = positionService;
        this.messagingTemplate = messagingTemplate;
        this.trainRepo = trainRepo;
        this.routeRepo = routeRepo;
        this.liveStateRepo = liveStateRepo;
        this.historyRepo = historyRepo;
        this.config = config;
    }

    public TrainStateMessage startTracking(Long trainId) {
        trackedTrains.add(trainId);
        updateCounters.put(trainId, new AtomicInteger(0));
        log.info("Started tracking train id={}", trainId);
        return processTrainUpdate(trainId);
    }

    public void stopTracking(Long trainId) {
        trackedTrains.remove(trainId);
        updateCounters.remove(trainId);
        log.info("Stopped tracking train id={}", trainId);
    }

    @Scheduled(fixedDelayString = "${app.provider.poll-interval-ms:15000}")
    public void pollAll() {
        for (Long trainId : trackedTrains) {
            try {
                processTrainUpdate(trainId);
            } catch (Exception e) {
                log.error("Error processing train {}: {}", trainId, e.getMessage(), e);
            }
        }
    }

    public TrainStateMessage processTrainUpdate(Long trainId) {
        Train train = trainRepo.findById(trainId).orElse(null);
        if (train == null) { stopTracking(trainId); return null; }

        Route route = routeRepo.findByTrain_TrainId(trainId).orElse(null);
        String geoJson = (route != null) ? route.getGeometryGeoJson() : null;

        if (!config.isSimulationEnabled() && (geoJson == null || geoJson.isBlank())) {
            String freshRoute = providerService.fetchRouteGeometry(train.getTrainNumber());
            if (freshRoute != null && freshRoute.length() > 50) {
                if (route == null) {
                    route = new Route();
                    route.setTrain(train);
                    route.setVersion(1);
                }
                route.setGeometryGeoJson(freshRoute);
                routeRepo.save(route);
                geoJson = freshRoute;
                log.info("Refreshed real railway route geometry for train {}", train.getTrainNumber());
            }
        }

        ProviderTelemetry telemetry = null;
        boolean isSimulated = config.isSimulationEnabled();

        if (isSimulated) {
            telemetry = simulationService.nextPosition(train.getTrainNumber(), geoJson);
        } else {
            telemetry = providerService.fetchLiveStatus(train.getTrainNumber());
            if (telemetry == null) {
                // If live API quota is exceeded (HTTP 429) or offline, fall back to smooth track simulation
                telemetry = simulationService.nextPosition(train.getTrainNumber(), geoJson);
                if (telemetry != null) {
                    telemetry.setStatus("RUNNING (ESTIMATED)");
                }
            }
        }

        TrainStateMessage msg = buildMessage(train, telemetry, geoJson, isSimulated);
        persistState(train, msg);
        maybeRecordHistory(train, msg);

        messagingTemplate.convertAndSend("/topic/train/" + trainId, msg);
        log.debug("Published update for train {} → stale={}", train.getTrainNumber(), msg.isStale());
        return msg;
    }

    private TrainStateMessage buildMessage(Train train, ProviderTelemetry t,
                                           String geoJson, boolean simulated) {
        TrainStateMessage msg = new TrainStateMessage();
        msg.setTrainId(train.getTrainId());
        msg.setTrainNumber(train.getTrainNumber());
        msg.setTrainName(train.getTrainName());
        if (train.getSourceStation() != null)
            msg.setSourceStation(train.getSourceStation().getStationName());
        if (train.getDestinationStation() != null)
            msg.setDestinationStation(train.getDestinationStation().getStationName());

        if (t == null) {
            // Provider unavailable — retain last known state as STALE
            liveStateRepo.findByTrain_TrainId(train.getTrainId()).ifPresent(ls -> {
                msg.setLatitude(ls.getLatitude());
                msg.setLongitude(ls.getLongitude());
                msg.setSpeedKmh(ls.getSpeedKmh());
                msg.setBearing(ls.getBearing());
                msg.setDelayMinutes(ls.getDelayMinutes());
                msg.setLastUpdated(ls.getLastUpdated());
            });
            msg.setStatus("STALE");
            msg.setStale(true);
            msg.setPositionSource("RETAINED");
            return msg;
        }

        // Validate coords
        Double lat = t.getLatitude(), lon = t.getLongitude();
        String positionSource = simulated ? "SIMULATED" : "PROVIDER";

        if (lat != null && lon != null
                && !Double.isNaN(lat) && !Double.isNaN(lon)
                && positionService.isValidCoordinate(lat, lon)
                && geoJson != null) {
            // Snap to route if needed
            double[] snapped = positionService.snapToRoute(
                    geoJson, lat, lon, config.getSnapToleranceMetres());
            if (snapped[0] != lat || snapped[1] != lon) {
                positionSource = simulated ? "SIMULATED_SNAPPED" : "SNAPPED";
            }
            lat = snapped[0];
            lon = snapped[1];
        } else if (t.getSegmentFromIndex() != null && t.getSegmentProgress() != null && geoJson != null) {
            // Interpolate from segment progress between points
            double[] interpolated = positionService.interpolateOnRoute(
                    geoJson, t.getSegmentFromIndex(), t.getSegmentProgress());
            if (interpolated != null) { lat = interpolated[0]; lon = interpolated[1]; }
            positionSource = "INTERPOLATED";
        } else if (t.getSegmentProgress() != null && geoJson != null) {
            // Interpolate from route-wide progress fraction
            double[] interpolated = positionService.interpolateAlongEntireRoute(geoJson, t.getSegmentProgress());
            if (interpolated != null) { lat = interpolated[0]; lon = interpolated[1]; }
            positionSource = "INTERPOLATED_PROGRESS";
        }

        msg.setLatitude(lat);
        msg.setLongitude(lon);
        msg.setSpeedKmh(t.getSpeedKmh());
        msg.setBearing(t.getBearing() != null ? positionService.normaliseBearing(t.getBearing()) : null);
        msg.setDelayMinutes(t.getDelayMinutes());
        msg.setStatus(t.getStatus() != null ? t.getStatus() : "UNKNOWN");
        msg.setLastUpdated(t.getTimestampEpochMs() != null
                ? Instant.ofEpochMilli(t.getTimestampEpochMs()) : Instant.now());
        msg.setPositionSource(positionSource);
        msg.setCurrentStation(t.getCurrentStation());
        msg.setNextHalt(t.getNextHalt());
        msg.setDistanceFromOriginKm(t.getDistanceFromOriginKm());
        msg.setTotalDistanceKm(t.getTotalDistanceKm());

        // Check staleness
        long ageMs = Instant.now().toEpochMilli()
                - (t.getTimestampEpochMs() != null ? t.getTimestampEpochMs() : Instant.now().toEpochMilli());
        msg.setStale(ageMs > config.getStaleThresholdMs());

        return msg;
    }

    private void persistState(Train train, TrainStateMessage msg) {
        if (msg.getLatitude() == null) return;
        LiveTrainState ls = liveStateRepo.findByTrain_TrainId(train.getTrainId())
                .orElseGet(() -> { LiveTrainState s = new LiveTrainState(); s.setTrain(train); return s; });
        ls.setLatitude(msg.getLatitude());
        ls.setLongitude(msg.getLongitude());
        ls.setSpeedKmh(msg.getSpeedKmh());
        ls.setBearing(msg.getBearing());
        ls.setDelayMinutes(msg.getDelayMinutes());
        ls.setStatus(msg.getStatus());
        ls.setLastUpdated(msg.getLastUpdated());
        ls.setPositionSource(msg.getPositionSource());
        liveStateRepo.save(ls);
    }

    private void maybeRecordHistory(Train train, TrainStateMessage msg) {
        if (config.getHistorySnapshotEvery() <= 0 || msg.getLatitude() == null) return;
        AtomicInteger counter = updateCounters.computeIfAbsent(
                train.getTrainId(), k -> new AtomicInteger(0));
        if (counter.incrementAndGet() % config.getHistorySnapshotEvery() == 0) {
            TrackingHistory h = new TrackingHistory();
            h.setTrain(train);
            h.setLatitude(msg.getLatitude());
            h.setLongitude(msg.getLongitude());
            h.setSpeedKmh(msg.getSpeedKmh());
            h.setTimestamp(msg.getLastUpdated() != null ? msg.getLastUpdated() : Instant.now());
            historyRepo.save(h);
        }
    }
}
