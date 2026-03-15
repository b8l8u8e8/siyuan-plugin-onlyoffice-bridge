FROM node:18-alpine
WORKDIR /app
COPY bridge.js ./
EXPOSE 6789
HEALTHCHECK --interval=30s --timeout=5s --retries=3 \
  CMD wget -qO- http://localhost:6789/health || exit 1
USER node
CMD ["node", "bridge.js"]
