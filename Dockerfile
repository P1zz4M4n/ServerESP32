# --- Fase 1: build delle dipendenze ---
# better-sqlite3 può richiedere compilazione nativa.
FROM node:22-slim AS build

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        python3 \
        make \
        g++ \
    && rm -rf /var/lib/apt/lists/*

COPY package*.json ./

RUN npm ci --omit=dev


# --- Fase 2: immagine finale ---
FROM node:22-slim

WORKDIR /app

ENV NODE_ENV=production \
    PORT=3000 \
    DB_PATH=/data/dipendenti.db

COPY --from=build /app/node_modules ./node_modules
COPY package*.json server.js ./

RUN mkdir -p /data \
    && chown node:node /data

USER node

EXPOSE 3000

CMD ["node", "server.js"]