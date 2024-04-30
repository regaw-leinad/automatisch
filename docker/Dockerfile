# syntax=docker/dockerfile:1
FROM node:18-alpine

ENV PORT 3000

RUN \
  apk --no-cache add --virtual build-dependencies python3 build-base git
RUN apk add --no-cache openssl dos2unix

WORKDIR /automatisch

# copy the app, note .dockerignore
COPY . /automatisch

RUN yarn

RUN cd packages/web && yarn build

RUN \
  rm -rf /usr/local/share/.cache/ && \
  apk del build-dependencies

COPY ./docker/entrypoint.sh /entrypoint.sh
COPY ./compose-entrypoint.sh /compose-entrypoint.sh
RUN dos2unix /compose-entrypoint.sh

EXPOSE 3000
ENTRYPOINT ["sh", "/compose-entrypoint.sh"]
