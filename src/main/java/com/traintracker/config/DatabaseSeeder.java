package com.traintracker.config;

import com.traintracker.model.Route;
import com.traintracker.model.Station;
import com.traintracker.model.Train;
import com.traintracker.repository.RouteRepository;
import com.traintracker.repository.StationRepository;
import com.traintracker.repository.TrainRepository;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.context.event.ApplicationReadyEvent;
import org.springframework.context.event.EventListener;
import org.springframework.stereotype.Component;

import java.util.Map;
import java.util.HashMap;
import java.io.BufferedReader;
import java.io.InputStreamReader;

@Component
public class DatabaseSeeder {

    private static final Logger log = LoggerFactory.getLogger(DatabaseSeeder.class);

    private final StationRepository stationRepo;
    private final TrainRepository trainRepo;
    private final RouteRepository routeRepo;
    private final org.springframework.core.io.ResourceLoader resourceLoader;

    public DatabaseSeeder(StationRepository stationRepo, TrainRepository trainRepo, RouteRepository routeRepo, org.springframework.core.io.ResourceLoader resourceLoader) {
        this.stationRepo = stationRepo;
        this.trainRepo = trainRepo;
        this.routeRepo = routeRepo;
        this.resourceLoader = resourceLoader;
    }


    @EventListener(ApplicationReadyEvent.class)
    public void seed() {
        if (stationRepo.count() > 0) {
            log.info("Database already seeded, skipping.");
            return;
        }

        log.info("Seeding initial railway stations, trains, and routes...");

        Map<String, Station> stations = new HashMap<>();
        stations.put("MMCT", createStation("MMCT", "Mumbai Central", 18.9696, 72.8197));
        stations.put("NDLS", createStation("NDLS", "New Delhi", 28.6428, 77.2197));
        stations.put("RKMP", createStation("RKMP", "Rani Kamalapati (Bhopal)", 23.2036, 77.4398));
        stations.put("HWH",  createStation("HWH",  "Howrah Junction", 22.5839, 88.3426));
        stations.put("SDAH", createStation("SDAH", "Sealdah (Kolkata)", 22.5675, 88.3710));
        stations.put("BKN",  createStation("BKN",  "Bikaner Junction", 28.0167, 73.3119));
        stations.put("TVC",  createStation("TVC",  "Thiruvananthapuram Central", 8.4875, 76.9525));
        stations.put("NZM",  createStation("NZM",  "Hazrat Nizamuddin", 28.5888, 77.2536));
        stations.put("VSG",  createStation("VSG",  "Vasco da Gama (Goa)", 15.3982, 73.8123));
        stations.put("LJN",  createStation("LJN",  "Lucknow Junction", 26.8322, 80.9200));
        stations.put("SBC",  createStation("SBC",  "KSR Bengaluru City", 12.9784, 77.5720));
        stations.put("SBIB", createStation("SBIB", "Sabarmati BG (Ahmedabad)", 23.0805, 72.5841));
        stations.put("MAS",  createStation("MAS",  "MGR Chennai Central", 13.0827, 80.2707));
        stations.put("CBE",  createStation("CBE",  "Coimbatore Junction", 11.0016, 76.9628));
        stations.put("CNB",  createStation("CNB",  "Kanpur Central", 26.4499, 80.3319));
        stations.put("BPL",  createStation("BPL",  "Bhopal Junction", 23.2599, 77.4126));
        stations.put("NGP",  createStation("NGP",  "Nagpur Junction", 21.1458, 79.0882));
        stations.put("AGC",  createStation("AGC",  "Agra Cantt", 27.1591, 78.0120));

        for (Station s : stations.values()) {
            stationRepo.save(s);
        }

        // 1. 12951 Mumbai Tejas Rajdhani Express (MMCT -> NDLS via Surat, Vadodara, Ratlam, Kota)
        Train t1 = trainRepo.save(createTrain("12951", "Mumbai Tejas Rajdhani Express", stations.get("MMCT"), stations.get("NDLS")));
        createRoute(t1, "{\"type\":\"LineString\",\"coordinates\":[[72.8197,18.9696],[72.8569,19.2288],[72.8407,21.2044],[73.1812,22.3107],[75.0441,23.3340],[75.8648,25.2138],[76.3686,25.9928],[77.6737,27.4924],[77.2197,28.6428]]}", 1386.0);

        // 2. 12002 Bhopal Shatabdi Express (NDLS -> RKMP via Mathura, Agra, Gwalior, Jhansi, Bhopal)
        Train t2 = trainRepo.save(createTrain("12002", "Bhopal Shatabdi Express", stations.get("NDLS"), stations.get("RKMP")));
        createRoute(t2, "{\"type\":\"LineString\",\"coordinates\":[[77.2197,28.6428],[77.6737,27.4924],[78.0120,27.1591],[78.1820,26.2183],[78.5685,25.4484],[77.4126,23.2599],[77.4398,23.2036]]}", 707.0);

        // 3. 12301 Howrah Rajdhani Express (HWH -> NDLS via Asansol, Dhanbad, Gaya, DDU, Prayagraj, Kanpur)
        Train t3 = trainRepo.save(createTrain("12301", "Howrah Rajdhani Express", stations.get("HWH"), stations.get("NDLS")));
        createRoute(t3, "{\"type\":\"LineString\",\"coordinates\":[[88.3426,22.5839],[86.9746,23.6889],[86.4304,23.7957],[85.0002,24.8028],[83.1167,25.2818],[81.8349,25.4447],[80.3319,26.4499],[77.2197,28.6428]]}", 1451.0);

        // 4. 12259 Sealdah Bikaner AC Duronto Express (SDAH -> BKN via Dhanbad, DDU, Kanpur, New Delhi, Churu)
        Train t4 = trainRepo.save(createTrain("12259", "Sealdah Bikaner AC Duronto Express", stations.get("SDAH"), stations.get("BKN")));
        createRoute(t4, "{\"type\":\"LineString\",\"coordinates\":[[88.3710,22.5675],[86.4304,23.7957],[83.1167,25.2818],[80.3319,26.4499],[77.2197,28.6428],[75.8167,28.4333],[75.3833,28.6333],[74.9667,28.3000],[73.3119,28.0167]]}", 1917.0);

        // 5. 12626 Kerala Superfast Express (NDLS -> TVC via Agra, Gwalior, Bhopal, Nagpur, Vijayawada, Katpadi, Coimbatore, Ernakulam)
        Train t5 = trainRepo.save(createTrain("12626", "Kerala Superfast Express", stations.get("NDLS"), stations.get("TVC")));
        createRoute(t5, "{\"type\":\"LineString\",\"coordinates\":[[77.2197,28.6428],[78.0120,27.1591],[78.1820,26.2183],[77.4126,23.2599],[79.0882,21.1458],[79.3333,19.8500],[80.6480,16.5062],[79.5167,13.6500],[79.1325,12.9249],[78.1460,11.6643],[76.9628,11.0016],[76.2999,9.9816],[76.9525,8.4875]]}", 3031.0);

        // 6. 12780 Goa Express (NZM -> VSG via Agra, Gwalior, Jhansi, Bhopal, Bhusaval, Manmad, Pune, Belagavi, Madgaon)
        Train t6 = trainRepo.save(createTrain("12780", "Goa Express", stations.get("NZM"), stations.get("VSG")));
        createRoute(t6, "{\"type\":\"LineString\",\"coordinates\":[[77.2536,28.5888],[78.0120,27.1591],[78.1820,26.2183],[78.5685,25.4484],[77.4126,23.2599],[75.7833,21.0500],[74.5333,20.2500],[73.8567,18.5204],[74.6500,16.8333],[74.5000,15.8500],[73.9862,15.2832],[73.8123,15.3982]]}", 2164.0);

        // 7. 12004 Lucknow Shatabdi Express (NDLS -> LJN via Ghaziabad, Aligarh, Tundla, Etawah, Kanpur)
        Train t7 = trainRepo.save(createTrain("12004", "Lucknow Shatabdi Express", stations.get("NDLS"), stations.get("LJN")));
        createRoute(t7, "{\"type\":\"LineString\",\"coordinates\":[[77.2197,28.6428],[77.4538,28.6692],[78.0800,27.8900],[78.2400,27.2100],[79.0300,26.7700],[80.3319,26.4499],[80.9200,26.8322]]}", 512.0);

        // 8. 22691 KSR Bengaluru Rajdhani Express (SBC -> NZM via Dharmavaram, Guntakal, Secunderabad, Kazipet, Nagpur, Bhopal, Jhansi)
        Train t8 = trainRepo.save(createTrain("22691", "KSR Bengaluru Rajdhani Express", stations.get("SBC"), stations.get("NZM")));
        createRoute(t8, "{\"type\":\"LineString\",\"coordinates\":[[77.5720,12.9784],[77.7167,14.4167],[77.3667,15.1667],[78.5034,17.4399],[79.5167,17.9833],[79.0882,21.1458],[77.4126,23.2599],[78.5685,25.4484],[78.0120,27.1591],[77.2536,28.5888]]}", 2367.0);

        // 9. 12957 Swarna Jayanti Rajdhani Express (SBIB -> NDLS via Mahesana, Palanpur, Abu Road, Ajmer, Jaipur, Gurgaon)
        Train t9 = trainRepo.save(createTrain("12957", "Swarna Jayanti Rajdhani Express", stations.get("SBIB"), stations.get("NDLS")));
        createRoute(t9, "{\"type\":\"LineString\",\"coordinates\":[[72.5841,23.0805],[72.3833,23.6000],[72.4333,24.1667],[72.7833,24.4833],[73.2333,25.2167],[74.6399,26.4499],[75.7873,26.9124],[77.0266,28.4595],[77.2197,28.6428]]}", 934.0);

        // 10. 12675 Kovai Superfast Express (MAS -> CBE via Arakkonam, Katpadi, Jolarpettai, Salem, Erode, Tiruppur)
        Train t10 = trainRepo.save(createTrain("12675", "Kovai Superfast Express", stations.get("MAS"), stations.get("CBE")));
        createRoute(t10, "{\"type\":\"LineString\",\"coordinates\":[[80.2707,13.0827],[79.6667,13.0833],[79.1325,12.9249],[78.5833,12.5667],[78.1460,11.6643],[77.7274,11.3410],[77.3411,11.1085],[76.9628,11.0016]]}", 496.0);

        log.info("Database seeding of all 10 verified Indian Railways trains, stations, and routes complete.");
        importAllTrainsFromCsv();
    }

    // ---------------------------------------------------------------------
    // CSV bulk‑import helper – reads /train-catalogue.csv from classpath.
    // ---------------------------------------------------------------------
    private void importAllTrainsFromCsv() {
        try (java.io.InputStream is = DatabaseSeeder.class.getResourceAsStream("/train-catalogue.csv")) {
            if (is == null) {
                log.warn("train-catalogue.csv not found on classpath.");
                return;
            }
            try (BufferedReader br = new BufferedReader(new InputStreamReader(is, java.nio.charset.StandardCharsets.UTF_8))) {
                String line;
                int added = 0;
                while ((line = br.readLine()) != null) {
                    line = line.trim();
                    if (line.isEmpty() || line.startsWith("#")) continue;
                    String[] parts = line.split(",");
                    if (parts.length < 2) continue;
                    String number = parts[0].trim();
                    String name   = parts[1].trim();
                    // Avoid duplicates – the three seeded trains already exist.
                    if (trainRepo.findByTrainNumber(number).isEmpty()) {
                        Train t = new Train();
                        t.setTrainNumber(number);
                        t.setTrainName(name);
                        t.setStatus("UNKNOWN");
                        trainRepo.save(t);
                        added++;
                    }
                }
                log.info("CSV train catalogue import finished – {} new trains added.", added);
            }
        } catch (Exception e) {
            log.warn("Failed to import train‑catalogue.csv: {}", e.getMessage());
        }
    }


    private Station createStation(String code, String name, double lat, double lon) {
        Station s = new Station();
        s.setStationCode(code);
        s.setStationName(name);
        s.setLatitude(lat);
        s.setLongitude(lon);
        return s;
    }

    private Train createTrain(String number, String name, Station source, Station dest) {
        Train t = new Train();
        t.setTrainNumber(number);
        t.setTrainName(name);
        t.setSourceStation(source);
        t.setDestinationStation(dest);
        t.setStatus("RUNNING");
        return t;
    }

    private void createRoute(Train train, String fallbackGeoJson, double distKm) {
        String finalGeoJson = fallbackGeoJson;
        try {
            java.nio.file.Path p = java.nio.file.Path.of("src", "main", "resources", "routes", train.getTrainNumber() + ".json");
            if (java.nio.file.Files.exists(p)) {
                String fileContent = java.nio.file.Files.readString(p, java.nio.charset.StandardCharsets.UTF_8);
                if (fileContent.trim().startsWith("{")) {
                    finalGeoJson = fileContent;
                    log.info("Loaded real physical railway track for train {} from disk", train.getTrainNumber());
                }
            } else {
                java.io.InputStream is = Thread.currentThread().getContextClassLoader().getResourceAsStream("routes/" + train.getTrainNumber() + ".json");
                if (is == null) {
                    is = DatabaseSeeder.class.getResourceAsStream("/routes/" + train.getTrainNumber() + ".json");
                }
                if (is != null) {
                    try (java.io.InputStream stream = is) {
                        String fileContent = new String(stream.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8);
                        if (fileContent.trim().startsWith("{")) {
                            finalGeoJson = fileContent;
                            log.info("Loaded real physical railway track for train {} from classpath", train.getTrainNumber());
                        }
                    }
                }
            }
        } catch (Exception e) {
            log.warn("Could not read route file for train {}: {}", train.getTrainNumber(), e.getMessage());
        }
        Route r = new Route();
        r.setTrain(train);
        r.setGeometryGeoJson(finalGeoJson);
        r.setDistanceKm(distKm);
        r.setVersion(1);
        routeRepo.save(r);
    }
}
