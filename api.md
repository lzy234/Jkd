## 登录
POST https://bapi.jkdsaas.com/b/pc/login HTTP/1.1
Host: bapi.jkdsaas.com
Connection: keep-alive
Content-Length: 737
sec-ch-ua: "Not.A/Brand";v="8", "Chromium";v="114"
Accept: application/json, text/plain, */*
sec-ch-ua-platform: "Windows"
Jkdauth: 
sec-ch-ua-mobile: ?0
User-Agent: Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) JKD-WSA-5/1.1.8 Chrome/114.0.5735.289 Electron/25.9.8 Safari/537.36
Content-Type: application/x-www-form-urlencoded;charset=UTF-8
Sec-Fetch-Site: same-site
Sec-Fetch-Mode: cors
Sec-Fetch-Dest: empty
Referer: https://b.jkdsaas.com/
Accept-Encoding: gzip, deflate, br
Accept-Language: zh-CN
Cookie: _pati=EodKno1HEgkkmlsteVW14BS9ZmdqJYDE; _pati_v=v2

un=HnOBDaTW17clBDMkxDp7C0ZlQLu36llhWtaHfnO2Ejp1JuqyGJ75WKdi3tjVD%2Fm9eHnPLVVmwpPQHAjgj8E%2FDYPbnhYWGFl2rSFyh2dWxLmlj1pp7f8dEkPKeCfuu1ysdc6BovvFs2k%2BgU%2FetDj8DC0jnscGAC8SXU%2FFNhqybcKJSfiP625OLvhXD7MIC7MuK0r%2BQ8NChlT9sudVD95c6uuvJFNm1feDJhLlOi2A4hBzlAG67kDECCncTTu0KKiRQHr0MAiGZB0bRqFaQk91sCW8NTVzb6gkCtaHuyIeNoSWJMib2dHI2AOq5GxTkDSi7t4uKUtJ8aWMmJ0YWmNngw%3D%3D&pw=ivbG3FKBhJSyL6sx7rKwMHp%2BAq0svpLbcIh5QToEDRznA4Pffv8uW2IWqpK2zXngytDE4L3jx9Pw1u3AatH7A2OW7hFK2%2FcsidTiBmt8AbcqB6U%2BA2zIVw4vD0vCHOYuqoMoZzGYAMWEMKl%2FDIm6j%2FAewmHIrTrVJTN1GMpDPIZnGzvHRYWZRHLRb5zqtvrT3XzzNuyx9%2BX1VXZug5iXz2JJeX23MyderbQtIyKyTmyURABl%2BJWUjxq%2B7KxZXULQ%2FPtpYm206qod%2FhCxaMVBEkFRE506sL0N52AorO8yJfviqrartNKMsoZ4gKmWZtttw%2FbkMWYIXisvzQjfFBsInw%3D%3D

## 获取派送中的订单（status=send）
GET https://bapi.jkdsaas.com/b/order/list?start_time=&end_time=&s_type=tel&bs_uuid=&sort_type=p_desc&status=send&k=&p=1&l=20&ch_uuid=&intf_token=GjZ9Bm&fileuuid=&bo_type=&is_printed=&order_type=wap HTTP/1.1
Host: bapi.jkdsaas.com
Connection: keep-alive
sec-ch-ua: "Not.A/Brand";v="8", "Chromium";v="114"
Accept: application/json, text/plain, */*
Jkdauth: eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzZXNzaW9uX2lkIjoiYXV0aHdub3QxdnBzeW5sdWZ5NGdsZ3IzNGk3aXVwdTNzdCIsInVzZXJfdXVpZCI6IlVTUlZTVlgxSVo5TURSWjhTRTI4SURFUUVEUE83TUxBIiwiYnVzZXJfdXVpZCI6IkJTVU1KS0dRU0VWS0E0UUJZVFFSTVJaTjEwRThOUURZIiwiaXNfYXBwIjpmYWxzZSwiZXhwX3RpbWUiOiIyMDI1LTEwLTA3IDE3OjUxOjExIn0.-ahZmeg0RdDmCvuO8Mna19H-VVUbbJYwEeYI4MkIMAA
sec-ch-ua-mobile: ?0
User-Agent: Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) JKD-WSA-5/1.1.8 Chrome/114.0.5735.289 Electron/25.9.8 Safari/537.36
sec-ch-ua-platform: "Windows"
Sec-Fetch-Site: same-site
Sec-Fetch-Mode: cors
Sec-Fetch-Dest: empty
Referer: https://b.jkdsaas.com/
Accept-Encoding: gzip, deflate, br
Accept-Language: zh-CN
Cookie: _pati=EodKno1HEgkkmlsteVW14BS9ZmdqJYDE; _pati_v=v2

## 修改备注
POST https://bapi.jkdsaas.com/b/order/set/buyer/memo HTTP/1.1
Host: bapi.jkdsaas.com
Connection: keep-alive
Content-Length: 43
sec-ch-ua: "Not.A/Brand";v="8", "Chromium";v="114"
Accept: application/json, text/plain, */*
sec-ch-ua-platform: "Windows"
Jkdauth: eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzZXNzaW9uX2lkIjoiYXV0aHdub3QxdnBzeW5sdWZ5NGdsZ3IzNGk3aXVwdTNzdCIsInVzZXJfdXVpZCI6IlVTUlZTVlgxSVo5TURSWjhTRTI4SURFUUVEUE83TUxBIiwiYnVzZXJfdXVpZCI6IkJTVU1KS0dRU0VWS0E0UUJZVFFSTVJaTjEwRThOUURZIiwiaXNfYXBwIjpmYWxzZSwiZXhwX3RpbWUiOiIyMDI1LTEwLTA3IDE3OjUxOjExIn0.-ahZmeg0RdDmCvuO8Mna19H-VVUbbJYwEeYI4MkIMAA
sec-ch-ua-mobile: ?0
User-Agent: Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) JKD-WSA-5/1.1.8 Chrome/114.0.5735.289 Electron/25.9.8 Safari/537.36
Content-Type: application/x-www-form-urlencoded;charset=UTF-8
Sec-Fetch-Site: same-site
Sec-Fetch-Mode: cors
Sec-Fetch-Dest: empty
Referer: https://b.jkdsaas.com/
Accept-Encoding: gzip, deflate, br
Accept-Language: zh-CN
Cookie: _pati=EodKno1HEgkkmlsteVW14BS9ZmdqJYDE; _pati_v=v2

order_uuid=251006-465531061920871&memo=test

## 修改派送人
POST https://bapi.jkdsaas.com/b/order/dispatch HTTP/1.1
Host: bapi.jkdsaas.com
Connection: keep-alive
Content-Length: 96
sec-ch-ua: "Not.A/Brand";v="8", "Chromium";v="114"
Accept: application/json, text/plain, */*
sec-ch-ua-platform: "Windows"
Jkdauth: eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJzZXNzaW9uX2lkIjoiYXV0aDB6ejdsdTdhZmYwZmxzZWRjdnpybGZkZ3ByeDlrMSIsInVzZXJfdXVpZCI6IlVTUlZTVlgxSVo5TURSWjhTRTI4SURFUUVEUE83TUxBIiwiYnVzZXJfdXVpZCI6IkJTVU1KS0dRU0VWS0E0UUJZVFFSTVJaTjEwRThOUURZIiwiaXNfYXBwIjpmYWxzZSwiZXhwX3RpbWUiOiIyMDI1LTEwLTEwIDE1OjIyOjA2In0.9sVN8hyB5mAuOf33pi2MZSzdGksdTOO8blXyeVHo6Eg
sec-ch-ua-mobile: ?0
User-Agent: Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) JKD-WSA-5/1.1.8 Chrome/114.0.5735.289 Electron/25.9.8 Safari/537.36
Content-Type: application/x-www-form-urlencoded;charset=UTF-8
Sec-Fetch-Site: same-site
Sec-Fetch-Mode: cors
Sec-Fetch-Dest: empty
Referer: https://b.jkdsaas.com/
Accept-Encoding: gzip, deflate, br
Accept-Language: zh-CN
Cookie: _pati=EodKno1HEgkkmlsteVW14BS9ZmdqJYDE; _pati_v=v2

order_uuid=251007-673458429652770&dm_uuid=BSUWOJ5UQ3QYPLLDLYXMEHR5R3CRCZIW&dm_salary=0&re_flag=1