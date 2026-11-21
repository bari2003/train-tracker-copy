# Multi-stage build for Spring Boot on Render
FROM maven:3.9.6-eclipse-temurin-17 AS build
WORKDIR /app

# Copy pom.xml and source code
COPY pom.xml .
COPY src ./src

# Build production JAR (skip unit tests for fast cloud build)
RUN mvn clean package -DskipTests

# Stage 2: Production JRE runtime
FROM eclipse-temurin:17-jre-jammy
WORKDIR /app

# Copy JAR from build stage
COPY --from=build /app/target/*.jar app.jar

# Tell Render both ports can receive traffic
EXPOSE 8080 10000

# Launch Spring Boot dynamically binding to Render's assigned port, all interfaces (0.0.0.0), and default h2 profile
ENTRYPOINT ["sh", "-c", "java -Dserver.port=${PORT:-8080} -Dserver.address=0.0.0.0 -Dspring.profiles.active=${SPRING_PROFILES_ACTIVE:-h2} -Djava.security.egd=file:/dev/./urandom -jar app.jar"]
