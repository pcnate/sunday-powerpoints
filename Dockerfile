
# syntax=docker/dockerfile:1

# ----------- Build Stage -----------
FROM node:24-alpine AS build

WORKDIR /app

# Copy package files and npmrc for dependency resolution
COPY package.json package-lock.json* .npmrc* ./

# Install dependencies (placeholder, no build step yet)
RUN npm install

# Copy all source files
COPY src ./src
COPY tsconfig.json jest.config.js ./

# ----------- Deploy Stage -----------
FROM node:24-alpine AS deploy

WORKDIR /app

# Copy only production dependencies
COPY package.json package-lock.json* .npmrc* ./
RUN npm install
# RUN npm install --omit=dev # have to leave the dev dependencies for ts-node to work

# Copy source files from build stage
COPY --from=build /app/src ./src
COPY --from=build /app/tsconfig.json ./

# Expose the port (adjust if needed)
EXPOSE 8080

# Set environment variables (can be overridden at runtime)
ENV NODE_ENV=production \
    TEMPLATE_FILE="Sunday Template.pptx" \
    EXT=pptx \
    TEMPLATE_DIRECTORY="/input" \
    OUTPUT_DIRECTORY="/output" \
    ROOT_PATH="%OneDriveConsumer%" \
    SONGS_DIRECTORY="/songs"

# Define mountable volumes for persistent files (adjust as needed)
VOLUME ["/input", "/output", "/songs"]

# Start the server with ts-node
CMD ["npx", "ts-node", "src/server.ts"]
