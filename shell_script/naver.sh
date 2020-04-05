curl -X GET 'polling.finance.naver.com/api/realtime.nhn?query=SERVICE_ITEM:'$(cat ./list.txt) | jq '.result.areas | .[0].datas'
