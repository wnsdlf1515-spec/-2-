FROM node:24-alpine
WORKDIR /app
COPY city-stage6/src /app/city-stage6/src
COPY city-stage6/index.html /app/city-stage6/index.html
COPY classroom /app/classroom
RUN node classroom/tools/build.cjs && mkdir -p /app/data && chown -R node:node /app
USER node
ENV NODE_ENV=production PORT=4190 DATA_FILE=/app/data/classroom.sqlite
EXPOSE 4190
CMD ["node", "classroom/server.cjs"]
