#!/bin/sh

CLASPJSON=$(cat <<-END
    {
        "scriptId": "$SCRIPT_ID",
        "rootDir": "./dist"
    }
END
)

# プロジェクト直下の .clasp.json を上書きし、意図した scriptId へ push する
echo "$CLASPJSON" > ./.clasp.json
