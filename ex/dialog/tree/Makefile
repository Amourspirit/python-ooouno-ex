BUILD_DIR  = tree_tabs
BUILD_ROOT_DIR = ../../../build
BUILD  = $(BUILD_ROOT_DIR)/$(BUILD_DIR)

help:
	@echo "Run make build to compile into 'build/$(BUILD_DIR)' folder."
	@echo "Run make build_ooodev to compile and includ OooDev into 'build/$(BUILD_DIR)' folder."
	@echo "Run make start to execute the dialog script locally"


.PHONY: build help, msg-short, msg-long, msg-warn, msg-error, msg-yes

build:
	oooscript compile --pyz-out --embed --config "$(PWD)/config.json" --build-dir "build/$(BUILD_DIR)"
	rm -f $(BUILD)/*.pyz
	rm -f $(BUILD)/*.py

build_ooodev:
	oooscript compile --pyz-out --embed --config "$(PWD)/config_ooodev.json" --build-dir "build/$(BUILD_DIR)"
	rm -f $(BUILD)/*.pyz
	rm -f $(BUILD)/*.py

start:
	python $(PWD)/start.py
