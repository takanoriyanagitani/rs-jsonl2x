#!/bin/sh

echo converting jsons to a sheet...
jq -c -n '[
	{timestamp:"2025-11-04T23:00:11.012345Z", severity:"INFO", status:200, body:"apt update done"},
	{timestamp:"2025-11-03T23:00:11.012345Z", severity:"WARN", status:500, body:"apt update failure"}
]' |
	jq -c '.[]' |
./jsonl2x \
	--output ./out.xlsx \
	--sheet Sheet1

echo
echo converting the sheet to jsonl...
x2jsonl \
	--input ./out.xlsx \
	--sheet Sheet1 
