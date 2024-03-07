build:
	npm run build && npm run ts-node

deploy:
	npm run build && npm run ts-node && clasp push && clasp open