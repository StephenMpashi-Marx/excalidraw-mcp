FROM node:22-slim

WORKDIR /app

COPY dist/ ./dist/

EXPOSE 3001

CMD ["node", "dist/index.js"]
