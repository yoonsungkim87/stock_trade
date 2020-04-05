curl -X GET 'https://api.binance.com/api/v3/ticker/24hr' | sed -r -e 's/\x22([0-9.]+)\x22/\1/g'
