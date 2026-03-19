FROM node:18-alpine
WORKDIR /app
COPY --chown=node:node bridge.js ./
RUN mkdir -p /app/cache && chown -R node:node /app
EXPOSE 27689
HEALTHCHECK --interval=30s --timeout=5s --retries=3 \
  CMD wget -qO- http://localhost:27689/health || exit 1
USER node
CMD ["node", "bridge.js"]
