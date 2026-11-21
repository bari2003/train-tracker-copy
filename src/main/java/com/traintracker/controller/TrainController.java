package com.traintracker.controller;

import com.traintracker.model.Route;
import com.traintracker.model.Train;
import com.traintracker.repository.RouteRepository;
import com.traintracker.service.TrackingService;
import com.traintracker.service.TrainManagementService;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.List;
import java.util.Map;

/**
 * REST API controller.
 *
 * GET  /api/trains?q=...         — search trains
 * GET  /api/trains/{id}          — get train details
 * GET  /api/trains/{id}/route    — get GeoJSON route
 * POST /api/trains/{id}/track    — start tracking
 * POST /api/trains/{id}/untrack  — stop tracking
 */
@RestController
@RequestMapping("/api/trains")
public class TrainController {

    private final TrainManagementService trainService;
    private final TrackingService trackingService;
    private final RouteRepository routeRepo;
    private final com.traintracker.service.LiveDataProviderService providerService;

    public TrainController(TrainManagementService trainService,
                           TrackingService trackingService,
                           RouteRepository routeRepo,
                           com.traintracker.service.LiveDataProviderService providerService) {
        this.trainService = trainService;
        this.trackingService = trackingService;
        this.routeRepo = routeRepo;
        this.providerService = providerService;
    }

    @GetMapping
    public ResponseEntity<List<Train>> search(@RequestParam(defaultValue = "") String q) {
        List<Train> trains = q.isBlank() ? trainService.getAll() : trainService.search(q);
        return ResponseEntity.ok(trains);
    }

    @GetMapping("/{id}")
    public ResponseEntity<Train> get(@PathVariable Long id) {
        Train t = trainService.getById(id);
        return t != null ? ResponseEntity.ok(t) : ResponseEntity.notFound().build();
    }

    @GetMapping("/{id}/route")
    public ResponseEntity<String> getRoute(@PathVariable Long id) {
        Train t = trainService.getById(id);
        if (t == null) return ResponseEntity.notFound().build();

        Route route = routeRepo.findByTrain_TrainId(id).orElse(null);
        if (route == null || route.getGeometryGeoJson() == null || route.getGeometryGeoJson().isBlank()) {
            String fresh = providerService.fetchRouteGeometry(t.getTrainNumber());
            if (fresh != null && fresh.length() > 50) {
                if (route == null) {
                    route = new Route();
                    route.setTrain(t);
                    route.setVersion(1);
                }
                route.setGeometryGeoJson(fresh);
                routeRepo.save(route);
            }
        }

        if (route != null && route.getGeometryGeoJson() != null) {
            return ResponseEntity.ok()
                    .header("Content-Type", "application/json")
                    .body(route.getGeometryGeoJson());
        }
        return ResponseEntity.notFound().build();
    }

    @GetMapping("/all-tracks")
    public ResponseEntity<List<Map<String, Object>>> getAllTracks() {
        List<Train> trains = trainService.getAll();
        List<Map<String, Object>> list = new java.util.ArrayList<>();
        for (Train t : trains) {
            Map<String, Object> item = new java.util.HashMap<>();
            item.put("trainId", t.getTrainId());
            item.put("trainNumber", t.getTrainNumber());
            item.put("trainName", t.getTrainName());
            item.put("status", t.getStatus());
            if (t.getSourceStation() != null) {
                item.put("sourceStation", t.getSourceStation().getStationName());
            }
            if (t.getDestinationStation() != null) {
                item.put("destinationStation", t.getDestinationStation().getStationName());
            }
            routeRepo.findByTrain_TrainId(t.getTrainId()).ifPresent(r -> {
                item.put("geometryGeoJson", r.getGeometryGeoJson());
            });
            list.add(item);
        }
        return ResponseEntity.ok(list);
    }

    @GetMapping("/live-api/{trainNumber}")
    public ResponseEntity<String> getLiveApiStatus(@PathVariable String trainNumber) {
        String raw = providerService.fetchRawLiveStatus(trainNumber);
        if (raw != null) {
            return ResponseEntity.ok()
                    .header("Content-Type", "application/json")
                    .body(raw);
        }
        return ResponseEntity.status(502)
                .header("Content-Type", "application/json")
                .body("{\"status\":\"ERROR\",\"message\":\"Unable to fetch live status for train " + trainNumber + " from RailCore API\"}");
    }

    @PostMapping("/{id}/track")
    public ResponseEntity<?> startTracking(@PathVariable Long id) {
        Train t = trainService.getById(id);
        if (t == null) return ResponseEntity.notFound().build();
        com.traintracker.dto.TrainStateMessage initialState = trackingService.startTracking(id);
        java.util.Map<String, Object> resp = new java.util.HashMap<>();
        resp.put("status", "tracking");
        resp.put("wsDestination", "/topic/train/" + id);
        if (initialState != null) {
            resp.put("initialState", initialState);
        }
        return ResponseEntity.ok(resp);
    }

    @PostMapping("/{id}/untrack")
    public ResponseEntity<Map<String, String>> stopTracking(@PathVariable Long id) {
        trackingService.stopTracking(id);
        return ResponseEntity.ok(Map.of("status", "stopped"));
    }
}
