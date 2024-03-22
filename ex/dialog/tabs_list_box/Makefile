help:
	@echo "Run make build to compile into 'build/sales_data' folder."
	@echo "Run make build_ooodev to compile and includ OooDev into 'build/sales_data' folder."
	@echo "Run make start to execute the dialog script locally"


.PHONY: build help, msg-short, msg-long, msg-warn, msg-error, msg-yes

build:
	oooscript compile --embed --config "$(PWD)/config.json" --embed-doc "$(PWD)/data/sales_data.ods" --build-dir "build/sales_data"

build_ooodev:
	oooscript compile --embed --config "$(PWD)/config_ooodev.json" --embed-doc "$(PWD)/data/sales_data.ods" --build-dir "build/sales_data"

start:
	python $(PWD)/start.py
