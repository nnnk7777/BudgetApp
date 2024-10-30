#!/bin/sh

CLASPJSON=$(cat <<-END
    {
        "scriptId": "$SCRIPT_ID",
        "rootDir": "./dist"
    }
END
)

echo $CLASPJSON > ~/.clasp.json
