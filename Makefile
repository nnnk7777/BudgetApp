.PHONY: all build copy deploy open

build:
	npm run build && npm run ts-node

build-and-push:
	make copy
	npm run build && npm run ts-node && clasp push --force

open:
	clasp open

copy:
	find dist -maxdepth 1 -type f -name '*.js' ! -name 'main.js' -delete
	find scripts -type f -name '*.js' ! -path 'scripts/dev/*' -exec cp {} dist/ \;
	cp src/config/categories.js dist/categories.js
	perl -0pi -e 's/^export default categories;\n?//m' dist/categories.js
	cp src/config/appsscript.json dist/appsscript.json

all:
	make copy
	make build-and-push
	make open
