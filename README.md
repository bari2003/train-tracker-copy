# Real-Time Train Movement Tracker

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/bari2003/train-tracker-copy)

A web-based real-time train tracking system built with **Spring Boot + WebSocket + Leaflet.js + H2/MySQL + RailCore API**.

---

## Project Structure

```
train-tracker/
├── pom.xml
├── sql/
│   └── schema.sql                        ← MySQL DDL + sample data
└── src/main/
    ├── java/com/traintracker/
    │   ├── TrainTrackerApplication.java
    │   ├── config/
    │   │   ├── AppConfig.java            ← All app properties
    │   │   └── WebSocketConfig.java      ← STOMP/SockJS setup
    │   ├── controller/
    │   │   └── TrainController.java      ← REST endpoints
    │   ├── dto/
    │   │   ├── ProviderTelemetry.java    ← Raw API response
    │   │   └── TrainStateMessage.java    ← WebSocket payload
    │   ├── model/                        ← JPA entities (3NF schema)
    │   │   ├── Train.java
    │   │   ├── Station.java
    │   │   ├── Route.java
    │   │   ├── RouteStop.java
    │   │   ├── LiveTrainState.java
    │   │   └── TrackingHistory.java
    │   ├── repository/                   ← Spring Data JPA
    │   └── service/
    │       ├── TrainManagementService.java
    │       ├── LiveDataProviderService.java  ← External API (RailRadar)
    │       ├── SimulationService.java        ← Fallback demo mode
    │       ├── TrackingService.java          ← Core pipeline + scheduler
    │       └── PositionProcessingService.java ← Validate/snap/interpolate
    └── resources/
        ├── application.properties
        └── static/
            ├── index.html
            ├── css/style.css
            └── js/app.js
```

---

## Quick Start

### 1. Prerequisites
- Java 17+
- Maven 3.8+
- MySQL 8+

### 2. Database Setup
```sql
-- Run the schema script
mysql -u root -p < sql/schema.sql
```

### 3. Configure application.properties
```properties
spring.datasource.url=jdbc:mysql://localhost:3306/train_tracker?...
spring.datasource.username=root
spring.datasource.password=YOUR_PASSWORD

# External API key (server-side only, never exposed to browser)
app.provider.api-key=YOUR_RAILRADAR_API_KEY

# Enable simulation mode (no live API needed)
app.simulation.enabled=true
```

### 4. Build & Run
```bash
mvn spring-boot:run
```

Open: **http://localhost:8080**

---

## REST API

| Method | Endpoint | Description |
|--------|----------|-------------|
| GET | `/api/trains?q=...` | Search trains by number/name |
| GET | `/api/trains/{id}` | Get train details |
| GET | `/api/trains/{id}/route` | Get GeoJSON route geometry |
| POST | `/api/trains/{id}/track` | Start live tracking |
| POST | `/api/trains/{id}/untrack` | Stop tracking |

## WebSocket

- **Endpoint**: `/ws` (SockJS)
- **Subscribe**: `/topic/train/{trainId}`
- **Message**: `TrainStateMessage` JSON

---

## Architecture

```
Browser (Leaflet.js + SockJS/STOMP)
        ↕ WebSocket /ws
Spring Boot Backend
  ├── REST Controller  ← train search, route, track/untrack
  ├── TrackingService  ← @Scheduled poller per tracked train
  │     ├── LiveDataProviderService  ← external API (RailRadar)
  │     ├── SimulationService        ← fallback demo
  │     └── PositionProcessingService ← validate/snap/interpolate
  └── MySQL (JPA/Hibernate)
        ├── trains, stations, routes, route_stops
        ├── live_train_states    ← latest position per train
        └── tracking_history     ← sampled snapshots
```

---

## Key Design Decisions (from SRS)

| Requirement | Implementation |
|-------------|----------------|
| API key never in browser JS | Stored in `application.properties`, used only in `LiveDataProviderService` |
| Stale data indication | `stale=true` flag in WebSocket message; indicator shown in UI |
| Simulation mode | `app.simulation.enabled=true` replays route waypoints |
| Route-level position snap | `PositionProcessingService.snapToRoute()` within configurable tolerance |
| Smooth marker animation | Frontend lerp interpolation over 800ms between WebSocket updates |
| Modular provider | `LiveDataProviderService` isolated behind interface — swap provider without touching tracking logic |
| 3NF schema | Trains, Stations, Routes, RouteStops, LiveTrainState, TrackingHistory separated |

---

## Simulation Mode

Set `app.simulation.enabled=true` — no external API key required. The system replays the stored GeoJSON route waypoints sequentially, looping at the end. Perfect for demo/testing.

---

## Technology Stack

| Layer | Technology |
|-------|-----------|
| Frontend | HTML5, CSS3, JavaScript, Leaflet.js |
| Real-time | WebSocket (SockJS + STOMP) |
| Backend | Java 17, Spring Boot 3.2 |
| Persistence | MySQL 8, Spring Data JPA / Hibernate |
| Build | Maven |
| External API | RailRadar (configurable) |
