# Install system dependencies for PDF processing (if not already installed)
if ! command -v pdftoppm &> /dev/null; then
    echo "Installing poppler-utils for PDF processing..."
    if command -v apt-get &> /dev/null; then
        # Ubuntu/Debian
        apt-get update && apt-get install -y poppler-utils
    elif command -v yum &> /dev/null; then
        # CentOS/RHEL
        yum install -y poppler-utils
    elif command -v apk &> /dev/null; then
        # Alpine
        apk add --no-cache poppler-utils
    else
        echo "Warning: Could not install poppler-utils automatically. PDF support may not work."
    fi
fi

mkdir -p ~/.streamlit/

echo "\
[general]\n\
email = \"samvitjatia9021@gmail.com\"\n\
" > ~/.streamlit/credentials.toml

echo "\
[server]\n\
headless = true\n\
enableCORS=false\n\
port = $PORT\n\
" > ~/.streamlit/config.toml
