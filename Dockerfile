FROM node:22.15.1-bullseye-slim
WORKDIR /app/
COPY ./ ./
RUN npm ci

CMD ["npm", "run", "dev"]
