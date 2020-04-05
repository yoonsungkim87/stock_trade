curl -X GET 'polling.finance.naver.com/api/realtime.nhn?query=SERVICE_ITEM:'$(cat /tmp/stock_trade/shell_script/list.txt) | jq '.result.areas | .[0].datas'
