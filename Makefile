.PHONY: all build deploy

build:
	npm run build && npm run ts-node

deploy:
	npm run build && npm run ts-node && clasp push && clasp open

copy:
	cp scripts/*.js dist/

all:
	make build
	make deploy