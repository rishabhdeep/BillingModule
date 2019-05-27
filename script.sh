#!/usr/bin/env bash
declare -a clients=(
    "merino"
    "cosma"
    "mahindra_net"
    "powerol"
    "rajasthan"
    "mahindraauto"
    "mahindra_tractor_outbound"
    "mahindra_tractor_inbound"
    "spicer"
    "siemens"
    "belden"
    "spares"
    "bulk"
    )

for i in "${clients[@]}"
do
    echo "$i"
    python3 main.py 1 1 2018 31 3 2019 data TRIPDAYS mahindra "$i"
    python3 main.py 1 4 2019 30 4 2019 data TRIPDAYS mahindra "$i"
done