package com.traintracker.service;

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.traintracker.dto.ProviderTelemetry;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * Simulation mode: replays pre-defined positions along the stored route geometry.
 * Activated when app.simulation.enabled=true.
 * Provides a repeatable demonstration path without depending on external live data.
 */
@Service
public class SimulationService {

    private static final Logger log = LoggerFactory.getLogger(SimulationService.class);
    private static final ObjectMapper mapper = new ObjectMapper();

    private final PositionProcessingService positionService;

    public SimulationService(PositionProcessingService positionService) {
        this.positionService = positionService;
    }

    /** Current waypoint index per train (loops when reaching end). */
    private final Map<String, AtomicInteger> positions = new ConcurrentHashMap<>();

    /**
     * Returns the next simulated telemetry position, stepping through the route waypoints.
     */
    public ProviderTelemetry nextPosition(String trainNumber, String geoJson) {
        ProviderTelemetry t = new ProviderTelemetry();
        t.setTrainNumber(trainNumber);
        t.setStatus("RUNNING");
        t.setTimestampEpochMs(System.currentTimeMillis());
        t.setSpeedKmh(78.0);
        t.setDelayMinutes(0);

        if (geoJson == null) {
            log.warn("Simulation: no route geometry for train {}", trainNumber);
            return t;
        }

        try {
            JsonNode root = mapper.readTree(geoJson);
            JsonNode coords = positionService.extractCoordinatesNode(root);
            if (coords == null || coords.isMissingNode() || coords.size() == 0) return t;

            AtomicInteger idx = positions.computeIfAbsent(trainNumber, k -> {
                int total = coords.size();
                if (total <= 4) return new AtomicInteger(0);
                int hash = Math.abs(trainNumber.hashCode());
                int initial = (int) (total * (0.25 + (hash % 45) / 100.0));
                return new AtomicInteger(Math.min(initial, total - 2));
            });
            int i = idx.getAndUpdate(v -> (v + 1) % coords.size());

            JsonNode pt = coords.get(i);
            t.setLongitude(pt.get(0).asDouble());
            t.setLatitude(pt.get(1).asDouble());
            t.setSegmentProgress((double) i / (double) coords.size());

            // Calculate bearing to next point
            if (i < coords.size() - 1) {
                JsonNode next = coords.get(i + 1);
                double bearing = calculateBearing(
                        pt.get(1).asDouble(), pt.get(0).asDouble(),
                        next.get(1).asDouble(), next.get(0).asDouble());
                t.setBearing(bearing);
            }
        } catch (Exception e) {
            log.error("Simulation step error for train {}: {}", trainNumber, e.getMessage());
        }
        return t;
    }

    private double calculateBearing(double lat1, double lon1, double lat2, double lon2) {
        double dLon = Math.toRadians(lon2 - lon1);
        double rlat1 = Math.toRadians(lat1), rlat2 = Math.toRadians(lat2);
        double x = Math.sin(dLon) * Math.cos(rlat2);
        double y = Math.cos(rlat1) * Math.sin(rlat2) - Math.sin(rlat1) * Math.cos(rlat2) * Math.cos(dLon);
        return (Math.toDegrees(Math.atan2(x, y)) + 360) % 360;
    }
}
