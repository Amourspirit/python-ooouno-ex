BUILD_DIR  = sales_grid
BUILD_ROOT_DIR = ../../../build
BUILD  = $(BUILD_ROOT_DIR)/$(BUILD_DIR)

help:
	@echo "Run make build to compile into 'build/$(BUILD_DIR)' folder."
	@echo "Run make build_ooodev to compile and includ OooDev into 'build/$(BUILD_DIR)' folder."
	@echo "Run make start to execute the dialog script locally"


.PHONY: build help

build:
	oooscript compile --embed --config "$(PWD)/config.json" --embed-doc "$(PWD)/data/sales_data.ods" --build-dir "build/$(BUILD_DIR)"
	rm -f $(BUILD)/*.py

build_ooodev:
	oooscript compile --pyz-out --embed --config "$(PWD)/config_ooodev.json" --embed-doc "$(PWD)/data/sales_data.ods" --build-dir "build/$(BUILD_DIR)"
	rm -f $(BUILD)/*.pyz
	rm -f $(BUILD)/*.py

start:
	python $(PWD)/start.