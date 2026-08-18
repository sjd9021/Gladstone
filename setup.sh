# PDF rendering needs no system packages: pypdfium2 ships PDFium as a
# prebuilt wheel. The poppler install that used to live here never ran
# on Render anyway - the service start command is 'streamlit run app.py',
# so this script is not executed there at all.

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
