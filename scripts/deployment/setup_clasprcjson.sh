#!/bin/sh

LOGIN=$(cat <<-END
    {
        "token": {
            "access_token": "$ACCESS_TOKEN",
            "refresh_token": "$REFRESH_TOKEN",
            "scope": "https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/drive.metadata.readonly https://www.googleapis.com/auth/script.projects https://www.googleapis.com/auth/script.deployments https://www.googleapis.com/auth/script.webapp.deploy https://www.googleapis.com/auth/logging.read https://www.googleapis.com/auth/userinfo.profile https://www.googleapis.com/auth/service.management https://www.googleapis.com/auth/cloud-platform openid https://www.googleapis.com/auth/userinfo.email",
            "token_type": "Bearer",
            "id_token": "$ID_TOKEN",
            "expiry_date": 1736786164865
        },
        "oauth2ClientSettings": {
            "clientId": "$CLIENT_ID",
            "clientSecret": "$CLIENT_SECRET",
            "redirectUri": "http://localhost"
        },
        "isLocalCreds": false,
        "tokens": {
            "default": {
            "client_id": "$CLIENT_ID",
            "client_secret": "$CLIENT_SECRET",
            "type": "authorized_user",
            "refresh_token": "$REFRESH_TOKEN_DEFAULT",
            "access_token": "$ACCESS_TOKEN_DEFAULT"
            }
        }
    }
END
)

echo $LOGIN > ~/.clasprc.json
