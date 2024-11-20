.PHONY: all build copy deploy open

build:
	npm run build && npm run ts-node

build-and-push:
	npm run build && npm run ts-node && clasp push --force

open:
	clasp open

copy:
	cp scripts/*.js dist/

all:
	make copy
	make build-and-deploy
	make open