# ============================================================
# Multi-stage build for Office LLM Harness Static File Server
# Serves PowerPoint/Word/Excel/Outlook add-in task pane UI files
# Uses Express to serve static files + dynamic manifest.xml
# ============================================================

# --- Stage 1: Build PowerPoint add-in static files ---
FROM node:20-alpine AS addin-builder
WORKDIR /build/addin

COPY src/office-addin/package.json src/office-addin/package-lock.json ./
RUN npm ci

COPY src/office-addin/webpack.config.js ./
COPY src/office-addin/tsconfig.json ./
COPY src/office-addin/assets/ ./assets/
COPY src/office-addin/src/ ./src/
COPY src/office-addin/manifest.xml ./

RUN npm run build

# --- Stage 2: Runtime image (Node.js + Express) ---
FROM node:20-alpine AS runtime
WORKDIR /app

# Copy package files and install dependencies
COPY server/package.json ./
RUN npm install --production

# Copy built static files from builder
COPY --from=addin-builder /build/addin/dist ./dist

# Copy dynamic manifest template (same content as source manifest.xml)
COPY --from=addin-builder /build/addin/manifest.xml ./manifest.xml

# Copy Express server that serves static files + dynamic manifest
COPY server/index.js ./index.js

EXPOSE 80

HEALTHCHECK --interval=30s --timeout=5s --start-period=5s --retries=3 \
    CMD wget -qO- http://localhost:80/health || exit 1

CMD ["node", "index.js"]
