package com.traintracker.service;

import com.traintracker.config.AppConfig;
import com.traintracker.dto.ProviderTelemetry;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.*;
import org.springframework.stereotype.Service;
import org.springframework.web.client.RestTemplate;

/**
 * Encapsulates communication with the external live-data provider (RailRadar).
 * Replace or extend this class to swap providers without changing tracking logic.
 */
@Service
public class LiveDataProviderService {

    private static final Logger log = LoggerFactory.getLogger(LiveDataProviderService.class);

    private final RestTemplate restTemplate;
    private final AppConfig config;
    private final ObjectMapper objectMapper = new ObjectMapper();

    public LiveDataProviderService(RestTemplate restTemplate, AppConfig config) {
        this.restTemplate = restTemplate;
        this.config = config;
    }

    private HttpHeaders createHeaders() {
        HttpHeaders headers = new HttpHeaders();
        String apiKey = config.getApiKey();
        String baseUrl = config.getProviderBaseUrl();
        if (apiKey != null && !apiKey.isBlank() && !apiKey.equals("YOUR_API_KEY_HERE")) {
            if (apiKey.startsWith("rk_") || baseUrl.contains("railcore")) {
                headers.set("X-RailCore-Key", apiKey);
            } else if (apiKey.startsWith("rwe1_") || baseUrl.contains("railway-engine") || baseUrl.contains("apkscope")) {
                headers.set("x-api-key", apiKey);
            } else {
                headers.set("Authorization", "Bearer " + apiKey);
            }
        }
        headers.setAccept(java.util.List.of(MediaType.APPLICATION_JSON));
        return headers;
    }

    private String buildLiveUrl(String trainNumber, String dateStr) {
        String baseUrl = config.getProviderBaseUrl();
        String apiKey = config.getApiKey();
        if (baseUrl.contains("railcore") || (apiKey != null && apiKey.startsWith("rk_"))) {
            return baseUrl.replaceAll("/+$", "") + "/v1/trains/" + trainNumber + "/live?date=" + dateStr;
        } else if (baseUrl.contains("railway-engine") || baseUrl.contains("apkscope")) {
            return baseUrl.replaceAll("/+$", "") + "/v1/train/" + trainNumber + "/status";
        } else if (baseUrl.contains("/v1") || baseUrl.contains("railradar.in")) {
            return baseUrl.replaceAll("/+$", "") + "/trains/" + trainNumber + "/live";
        } else {
            String sep = baseUrl.contains("?") ? "&" : "?";
            return baseUrl + "/api/live-train-status" + sep + "train=" + trainNumber;
        }
    }

    public String fetchRawLiveStatus(String trainNumber) {
        try {
            java.time.LocalDate today = java.time.LocalDate.now(java.time.ZoneId.of("Asia/Kolkata"));
            String url = buildLiveUrl(trainNumber, today.toString());
            HttpHeaders headers = createHeaders();

            ResponseEntity<String> response = restTemplate.exchange(
                    url,
                    HttpMethod.GET,
                    new HttpEntity<>(headers),
                    String.class
            );

            if (response.getStatusCode().is2xxSuccessful() && response.getBody() != null) {
                String body = response.getBody();
                // If today's status is NOT_STARTED or progress is 0, check if yesterday was running
                try {
                    JsonNode root = objectMapper.readTree(body);
                    JsonNode dataNode = root.has("data") ? root.path("data") : root;
                    String status = dataNode.path("status").asText("");
                    double progress = dataNode.path("progress_percent").asDouble(0.0);
                    if ("NOT_STARTED".equalsIgnoreCase(status) || progress == 0.0) {
                        java.time.LocalDate yesterday = today.minusDays(1);
                        String yesterdayUrl = buildLiveUrl(trainNumber, yesterday.toString());
                        ResponseEntity<String> yResp = restTemplate.exchange(
                                yesterdayUrl,
                                HttpMethod.GET,
                                new HttpEntity<>(headers),
                                String.class
                        );
                        if (yResp.getStatusCode().is2xxSuccessful() && yResp.getBody() != null) {
                            JsonNode yRoot = objectMapper.readTree(yResp.getBody());
                            JsonNode yData = yRoot.has("data") ? yRoot.path("data") : yRoot;
                            String yStatus = yData.path("status").asText("");
                            if ("RUNNING".equalsIgnoreCase(yStatus) || yData.path("progress_percent").asDouble(0) > 0) {
                                return yResp.getBody();
                            }
                        }
                    }
                } catch (Exception ignored) {}
                return body;
            }
        } catch (Exception e) {
            log.warn("Raw provider fetch failed for train {}: {}", trainNumber, e.getMessage());
        }
        return null;
    }

    /**
     * Fetches live running status for a given train number.
     * Returns null if provider is unavailable or data is unusable.
     */
    public ProviderTelemetry fetchLiveStatus(String trainNumber) {
        try {
            java.time.LocalDate today = java.time.LocalDate.now(java.time.ZoneId.of("Asia/Kolkata"));
            String url = buildLiveUrl(trainNumber, today.toString());
            HttpHeaders headers = createHeaders();

            ResponseEntity<String> response = restTemplate.exchange(
                    url,
                    HttpMethod.GET,
                    new HttpEntity<>(headers),
                    String.class
            );

            if (response.getStatusCode().is2xxSuccessful() && response.getBody() != null) {
                ProviderTelemetry t = parseTelemetry(response.getBody(), trainNumber);
                if (t != null && ("NOT_STARTED".equalsIgnoreCase(t.getStatus()) || t.getSegmentProgress() == null || t.getSegmentProgress() == 0.0)) {
                    // Try checking yesterday's trip for overnight running trains
                    try {
                        java.time.LocalDate yesterday = today.minusDays(1);
                        String yesterdayUrl = buildLiveUrl(trainNumber, yesterday.toString());
                        ResponseEntity<String> yResp = restTemplate.exchange(
                                yesterdayUrl,
                                HttpMethod.GET,
                                new HttpEntity<>(headers),
                                String.class
                        );
                        if (yResp.getStatusCode().is2xxSuccessful() && yResp.getBody() != null) {
                            ProviderTelemetry yTel = parseTelemetry(yResp.getBody(), trainNumber);
                            if (yTel != null && ("RUNNING".equalsIgnoreCase(yTel.getStatus()) || (yTel.getSegmentProgress() != null && yTel.getSegmentProgress() > 0))) {
                                return yTel;
                            }
                        }
                    } catch (Exception ignored) {}
                }
                return t;
            }
        } catch (Exception e) {
            log.warn("Provider fetch failed for train {}: {}", trainNumber, e.getMessage());
        }
        return null;
    }

    /**
     * Fetches GeoJSON route geometry for a given train number.
     * Returns null on failure.
     */
    public String fetchRouteGeometry(String trainNumber) {
        try {
            String baseUrl = config.getProviderBaseUrl();
            String apiKey = config.getApiKey();
            String url;
            if (baseUrl.contains("railcore") || (apiKey != null && apiKey.startsWith("rk_"))) {
                url = baseUrl.replaceAll("/+$", "") + "/v1/trains/" + trainNumber + "/route";
            } else if (baseUrl.contains("railway-engine") || baseUrl.contains("apkscope")) {
                url = baseUrl.replaceAll("/+$", "") + "/v1/train/" + trainNumber + "/route";
            } else if (baseUrl.contains("/v1") || baseUrl.contains("railradar.in")) {
                url = baseUrl.replaceAll("/+$", "") + "/trains/" + trainNumber + "/route?format=geojson&stops=true";
            } else {
                url = baseUrl + "/api/train-route-geometry?train=" + trainNumber;
            }

            HttpHeaders headers = createHeaders();

            ResponseEntity<String> response = restTemplate.exchange(
                    url,
                    HttpMethod.GET,
                    new HttpEntity<>(headers),
                    String.class
            );

            if (response.getStatusCode().is2xxSuccessful() && response.getBody() != null) {
                JsonNode root = objectMapper.readTree(response.getBody());
                JsonNode data = root.has("data") ? root.path("data") : root;

                // RailCore route format with stops array
                if (data.has("stops") && data.path("stops").isArray()) {
                    com.fasterxml.jackson.databind.node.ObjectNode fc = objectMapper.createObjectNode();
                    fc.put("type", "FeatureCollection");
                    com.fasterxml.jackson.databind.node.ArrayNode features = fc.putArray("features");

                    // 1. LineString feature connecting all stops/waypoints
                    com.fasterxml.jackson.databind.node.ObjectNode lineFeature = objectMapper.createObjectNode();
                    lineFeature.put("type", "Feature");
                    com.fasterxml.jackson.databind.node.ObjectNode lineGeom = lineFeature.putObject("geometry");
                    lineGeom.put("type", "LineString");
                    com.fasterxml.jackson.databind.node.ArrayNode lineCoords = lineGeom.putArray("coordinates");

                    for (JsonNode stop : data.path("stops")) {
                        double lat = stop.path("latitude").asDouble(0.0);
                        double lng = stop.path("longitude").asDouble(0.0);
                        if (lat != 0.0 && lng != 0.0) {
                            com.fasterxml.jackson.databind.node.ArrayNode pt = lineCoords.addArray();
                            pt.add(lng);
                            pt.add(lat);

                            // Point feature for stations
                            boolean isStop = stop.path("is_stop").asBoolean(true);
                            if (isStop) {
                                com.fasterxml.jackson.databind.node.ObjectNode pf = objectMapper.createObjectNode();
                                pf.put("type", "Feature");
                                com.fasterxml.jackson.databind.node.ObjectNode geom = pf.putObject("geometry");
                                geom.put("type", "Point");
                                com.fasterxml.jackson.databind.node.ArrayNode coords = geom.putArray("coordinates");
                                coords.add(lng);
                                coords.add(lat);
                                com.fasterxml.jackson.databind.node.ObjectNode props = pf.putObject("properties");
                                props.put("stationName", stop.path("station_name").asText("Station"));
                                props.put("stationCode", stop.path("station_code").asText(""));
                                features.add(pf);
                            }
                        }
                    }

                    // Add lineFeature first
                    if (lineCoords.size() > 0) {
                        features.insert(0, lineFeature);
                        return objectMapper.writeValueAsString(fc);
                    }
                }

                if (data.has("geojson") && data.has("stops") && data.path("stops").isArray()) {
                    com.fasterxml.jackson.databind.node.ObjectNode fc = objectMapper.createObjectNode();
                    fc.put("type", "FeatureCollection");
                    com.fasterxml.jackson.databind.node.ArrayNode features = fc.putArray("features");
                    features.add(data.path("geojson"));
                    for (JsonNode stop : data.path("stops")) {
                        double sLat = stop.path("lat").asDouble();
                        double sLng = stop.path("lng").asDouble();
                        if (sLat != 0 && sLng != 0) {
                            com.fasterxml.jackson.databind.node.ObjectNode pf = objectMapper.createObjectNode();
                            pf.put("type", "Feature");
                            com.fasterxml.jackson.databind.node.ObjectNode geom = pf.putObject("geometry");
                            geom.put("type", "Point");
                            com.fasterxml.jackson.databind.node.ArrayNode coords = geom.putArray("coordinates");
                            coords.add(sLng);
                            coords.add(sLat);
                            com.fasterxml.jackson.databind.node.ObjectNode props = pf.putObject("properties");
                            props.put("stationName", stop.path("name").asText("Station"));
                            props.put("stationCode", stop.path("code").asText(""));
                            features.add(pf);
                        }
                    }
                    return objectMapper.writeValueAsString(fc);
                } else if (data.has("geojson")) {
                    return objectMapper.writeValueAsString(data.path("geojson"));
                }
                return response.getBody();
            }
        } catch (Exception e) {
            log.warn("Route geometry fetch failed for train {}: {}", trainNumber, e.getMessage());
        }
        return null;
    }

    private ProviderTelemetry parseTelemetry(String json, String trainNumber) {
        try {
            JsonNode root = objectMapper.readTree(json);
            JsonNode dataNode = root.has("data") ? root.path("data") : root;
            JsonNode locNode = dataNode.has("currentLocation") ? dataNode.path("currentLocation") : dataNode;

            ProviderTelemetry t = new ProviderTelemetry();
            t.setTrainNumber(dataNode.path("train_number").asText(trainNumber));

            // 1. Latitude & Longitude (supports RailCore root fields, currentLocation, pos, etc.)
            if (dataNode.has("latitude") && !dataNode.path("latitude").isNull() && dataNode.path("latitude").asDouble() != 0.0) {
                t.setLatitude(dataNode.path("latitude").asDouble());
            } else if (locNode.has("latitude")) {
                t.setLatitude(locNode.path("latitude").asDouble(Double.NaN));
            } else if (locNode.has("lat")) {
                t.setLatitude(locNode.path("lat").asDouble(Double.NaN));
            }

            if (dataNode.has("longitude") && !dataNode.path("longitude").isNull() && dataNode.path("longitude").asDouble() != 0.0) {
                t.setLongitude(dataNode.path("longitude").asDouble());
            } else if (locNode.has("longitude")) {
                t.setLongitude(locNode.path("longitude").asDouble(Double.NaN));
            } else if (locNode.has("lng")) {
                t.setLongitude(locNode.path("lng").asDouble(Double.NaN));
            }

            if (dataNode.has("position") && dataNode.path("position").isObject()) {
                JsonNode pos = dataNode.path("position");
                if (pos.has("latitude") && !pos.path("latitude").isNull() && !Double.isNaN(pos.path("latitude").asDouble())) {
                    double lat = pos.path("latitude").asDouble();
                    double lon = pos.path("longitude").asDouble();
                    if (lat != 0.0 && lon != 0.0) {
                        t.setLatitude(lat);
                        t.setLongitude(lon);
                    }
                }
            }

            // 2. Speed (supports speed_kmph, speedKmph, speed)
            if (dataNode.has("speed_kmph") && !dataNode.path("speed_kmph").isNull()) {
                t.setSpeedKmh(dataNode.path("speed_kmph").asDouble());
            } else if (dataNode.has("speedKmph") && !dataNode.path("speedKmph").isNull()) {
                t.setSpeedKmh(dataNode.path("speedKmph").asDouble());
            } else if (locNode.has("speedKmh")) {
                t.setSpeedKmh(locNode.path("speedKmh").asDouble());
            } else if (locNode.has("speed")) {
                t.setSpeedKmh(locNode.path("speed").asDouble());
            }
            if (t.getSpeedKmh() == null || t.getSpeedKmh() <= 0.0) {
                String st = dataNode.path("status").asText("");
                t.setSpeedKmh("RUNNING".equalsIgnoreCase(st) ? 75.0 : 0.0);
            }

            if (locNode.has("bearingDegrees")) {
                t.setBearing(locNode.path("bearingDegrees").asDouble());
            } else if (locNode.has("bearing")) {
                t.setBearing(locNode.path("bearing").asDouble());
            }

            // 3. Current Station
            if (dataNode.has("current_station_name")) {
                String stn = dataNode.path("current_station_name").asText("");
                String code = dataNode.path("current_station_code").asText("");
                t.setCurrentStation(stn + (code.isBlank() ? "" : " (" + code + ")"));
            } else if (dataNode.has("currentStation") && dataNode.path("currentStation").isObject()) {
                JsonNode cs = dataNode.path("currentStation");
                String stn = cs.path("name").asText("");
                String code = cs.path("code").asText("");
                t.setCurrentStation(stn + (code.isBlank() ? "" : " (" + code + ")"));
            } else if (locNode.has("stationName") || locNode.has("stationCode")) {
                String stn = locNode.path("stationName").asText("");
                String code = locNode.path("stationCode").asText("");
                String stStatus = locNode.path("status").asText("");
                String prefix = stStatus.isBlank() ? "" : (stStatus.substring(0, 1).toUpperCase() + stStatus.substring(1) + " ");
                t.setCurrentStation(prefix + stn + (code.isBlank() ? "" : " (" + code + ")"));
            }

            // 4. Next Halt / Stop
            if (dataNode.has("next_stop") && dataNode.path("next_stop").isObject()) {
                JsonNode ns = dataNode.path("next_stop");
                String stn = ns.path("station_name").asText("");
                String code = ns.path("station_code").asText("");
                t.setNextHalt(stn + (code.isBlank() ? "" : " (" + code + ")"));
            } else if (dataNode.has("next_station_code") && !dataNode.path("next_station_code").asText("").isBlank()) {
                t.setNextHalt(dataNode.path("next_station_code").asText(""));
            } else if (dataNode.has("nextStation") && dataNode.path("nextStation").isObject()) {
                JsonNode ns = dataNode.path("nextStation");
                String stn = ns.path("name").asText("");
                String code = ns.path("code").asText("");
                t.setNextHalt(stn + (code.isBlank() ? "" : " (" + code + ")"));
            } else if (dataNode.has("nextHalt")) {
                JsonNode nh = dataNode.path("nextHalt");
                t.setNextHalt(nh.path("stationName").asText("") + " (" + nh.path("stationCode").asText("") + ")");
            }

            // 5. Route Progress
            if (dataNode.has("progress_percent")) {
                double pct = dataNode.path("progress_percent").asDouble();
                if (pct >= 0) {
                    t.setSegmentProgress(pct / 100.0);
                }
            } else if (dataNode.has("routeProgress") && dataNode.path("routeProgress").has("percent")) {
                double pct = dataNode.path("routeProgress").path("percent").asDouble();
                if (pct >= 0) {
                    t.setSegmentProgress(pct / 100.0);
                }
            } else if (locNode.has("distanceFromOriginKm") && dataNode.path("train").has("distance")) {
                double distOrigin = locNode.path("distanceFromOriginKm").asDouble();
                double totalDist = dataNode.path("train").path("distance").asDouble();
                t.setDistanceFromOriginKm(distOrigin);
                t.setTotalDistanceKm(totalDist);
                if (totalDist > 0 && distOrigin > 0) {
                    t.setSegmentProgress(distOrigin / totalDist);
                } else if (distOrigin == 0.0) {
                    t.setSegmentProgress(0.0);
                }
            } else if (locNode.has("segmentProgress")) {
                double sp = locNode.path("segmentProgress").asDouble();
                if (sp >= 0) {
                    t.setSegmentProgress(sp);
                }
            }

            // 6. Distances
            if (dataNode.has("distance_covered_km")) {
                t.setDistanceFromOriginKm(dataNode.path("distance_covered_km").asDouble());
            }
            if (dataNode.has("total_distance_km")) {
                t.setTotalDistanceKm(dataNode.path("total_distance_km").asDouble());
            }

            // 7. Delays
            if (dataNode.has("next_stop") && dataNode.path("next_stop").has("delay_arrival_minutes")) {
                t.setDelayMinutes(dataNode.path("next_stop").path("delay_arrival_minutes").asInt());
            } else if (dataNode.has("delayMinutes")) {
                t.setDelayMinutes(dataNode.path("delayMinutes").asInt());
            } else if (dataNode.path("delay").has("minutes")) {
                t.setDelayMinutes(dataNode.path("delay").path("minutes").asInt());
            } else if (dataNode.has("delay")) {
                t.setDelayMinutes(dataNode.path("delay").asInt());
            } else if (dataNode.has("status_text")) {
                String stText = dataNode.path("status_text").asText("");
                java.util.regex.Matcher m = java.util.regex.Pattern.compile("(\\d+)\\s+min").matcher(stText);
                if (m.find()) {
                    try {
                        t.setDelayMinutes(Integer.parseInt(m.group(1)));
                    } catch (Exception ignored) {}
                }
            }

            // 8. Status
            if (dataNode.has("status")) {
                t.setStatus(dataNode.path("status").asText("UNKNOWN").toUpperCase());
            } else {
                t.setStatus("RUNNING");
            }

            long timestamp = System.currentTimeMillis();
            if (dataNode.has("observedAt")) {
                try {
                    timestamp = java.time.OffsetDateTime.parse(dataNode.path("observedAt").asText()).toInstant().toEpochMilli();
                } catch (Exception ignored) {}
            } else if (root.has("meta") && root.path("meta").has("timestamp")) {
                try {
                    timestamp = java.time.Instant.parse(root.path("meta").path("timestamp").asText()).toEpochMilli();
                } catch (Exception ignored) {}
            } else if (dataNode.has("timestamp")) {
                timestamp = dataNode.path("timestamp").asLong(timestamp);
            }
            t.setTimestampEpochMs(timestamp);

            return t;
        } catch (Exception e) {
            log.error("Failed to parse telemetry JSON: {}", e.getMessage());
            return null;
        }
    }
}
