#!/bin/bash
# Deploy DPV Flash Update Generator to S3
# Usage: ./deploy-s3.sh <bucket-name> [cloudfront-distribution-id]

set -e

BUCKET=$1
DIST_ID=$2

if [ -z "$BUCKET" ]; then
    echo "Usage: ./deploy-s3.sh <bucket-name> [cloudfront-distribution-id]"
    exit 1
fi

echo "⚡ Deploying DPV Flash Update Generator to s3://$BUCKET/"

# Sync all files with appropriate content types
aws s3 sync . "s3://$BUCKET/" \
    --exclude "*.sh" \
    --exclude "README.md" \
    --exclude ".nojekyll" \
    --exclude ".git/*" \
    --cache-control "max-age=3600"

# Set correct content types
aws s3 cp "s3://$BUCKET/index.html" "s3://$BUCKET/index.html" \
    --content-type "text/html" \
    --cache-control "no-cache" \
    --metadata-directive REPLACE

aws s3 cp "s3://$BUCKET/styles.css" "s3://$BUCKET/styles.css" \
    --content-type "text/css" \
    --metadata-directive REPLACE

aws s3 cp "s3://$BUCKET/app.js" "s3://$BUCKET/app.js" \
    --content-type "application/javascript" \
    --metadata-directive REPLACE

echo "✅ Files uploaded to s3://$BUCKET/"

# Invalidate CloudFront if distribution ID provided
if [ -n "$DIST_ID" ]; then
    echo "🔄 Invalidating CloudFront cache..."
    aws cloudfront create-invalidation \
        --distribution-id "$DIST_ID" \
        --paths "/*" > /dev/null
    echo "✅ CloudFront invalidation started"
fi

echo ""
echo "🎉 Deploy complete!"
