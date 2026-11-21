package com.traintracker.service;

import com.traintracker.dto.ProviderTelemetry;
import com.traintracker.model.Route;
import com.traintracker.model.Train;
import com.traintracker.repository.RouteRepository;
import com.traintracker.repository.TrainRepository;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.context.annotation.Lazy;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.List;

/** Manages train reference data and search. */
@Service
public class TrainManagementService {

    private static final Logger log = LoggerFactory.getLogger(TrainManagementService.class);

    private final TrainRepository trainRepo;
    private final RouteRepository routeRepo;
    private final LiveDataProviderService providerService;

    public TrainManagementService(TrainRepository trainRepo,
                                  RouteRepository routeRepo,
                                  @Lazy LiveDataProviderService providerService) {
        this.trainRepo = trainRepo;
        this.routeRepo = routeRepo;
        this.providerService = providerService;
    }

    public List<Train> search(String query) {
        return trainRepo.searchTrains(query);
    }

    @Transactional
    public Train fetchAndCreateFromProvider(String trainNumber) {
        try {
            ProviderTelemetry tel = providerService.fetchLiveStatus(trainNumber);
            if (tel != null) { // Accept provider data even without live coordinates
                Train train = new Train();
                train.setTrainNumber(trainNumber);
                train.setTrainName("Train " + trainNumber);
                train.setStatus(tel.getStatus() != null ? tel.getStatus() : "RUNNING");
                train = trainRepo.save(train);

                String routeGeoJson = providerService.fetchRouteGeometry(trainNumber);
                if (routeGeoJson != null && !routeGeoJson.isBlank()) {
                    Route route = new Route();
                    route.setTrain(train);
                    route.setGeometryGeoJson(routeGeoJson);
                    route.setDistanceKm(100.0);
                    route.setVersion(1);
                    routeRepo.save(route);
                }
                log.info("Dynamically imported train {} from live provider", trainNumber);
                return train;
            }
        } catch (Exception e) {
            log.warn("Could not dynamically import train {}: {}", trainNumber, e.getMessage());
        }
        return null;
    }

    public Train getByNumber(String trainNumber) {
        return trainRepo.findByTrainNumber(trainNumber).orElse(null);
    }

    public Train getById(Long id) {
        return trainRepo.findById(id).orElse(null);
    }

    public List<Train> getAll() {
        return trainRepo.findAll();
    }
}
