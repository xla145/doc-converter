# Use an official OpenJDK runtime as a parent image
FROM openjdk:8-jdk-alpine

# Set the working directory in the container
WORKDIR /app

# Copy the current directory contents into the container at /app
COPY . .

# Package the application with Maven
RUN ./mvnw package

# Run the application
CMD ["java", "-jar", "target/doc-converter-0.0.1-SNAPSHOT.jar"] 