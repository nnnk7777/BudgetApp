.PHONY: all build copy deploy open

build:
	npm run build && npm run ts-node

build-and-deploy:
	npm run build && npm run ts-node && clasp push

open:
	clasp open

copy:
	cp scripts/*.js dist/

all:
	make copy
	make build
	make deploy
	make open