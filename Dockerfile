FROM node:24-alpine AS build-env
WORKDIR /app
ENV NODE_ENV=production

COPY package*.json ./
RUN --mount=type=cache,target=/root/.npm yarn install --production --ignore-scripts

COPY . .
RUN mkdir -p /app/uploads && chown -R 65532:65532 /app/uploads
FROM gcr.io/distroless/nodejs24-debian12:nonroot

COPY --from=build-env /app /app
WORKDIR /app
EXPOSE 3001

CMD ["index.js"]