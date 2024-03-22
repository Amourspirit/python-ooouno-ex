help:
	@echo "Run make build to compile into 'build/tree_tabs' folder."
	@echo "Run make build_ooodev to compile and includ OooDev into 'build/tree_tabs' folder."
	@echo "Run make start to execute the dialog script locally"


.PHONY: build help, msg-short, msg-long, msg-warn, msg-error, msg-yes

build:
	oooscript compile --embed --config "$(PWD)/config.json" --build-dir "build/tree_tabs"

build_ooodev:
	oooscript compile --embed --config "$(PWD)/config_ooodev.json" --build-dir "build/tree_tabs"

start:
	python $(PWD)/start.py
