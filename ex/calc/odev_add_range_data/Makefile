BUILD_DIR  = add_range_data
BUILD_ROOT_DIR = ../../../build
BUILD  = $(BUILD_ROOT_DIR)/$(BUILD_DIR)

help:
	@echo "Run make build to compile into 'build/$(BUILD_DIR)' folder."
	@echo "Run make macro to compile into 'macro/odev_add_range.py' file."
	@echo "Run make run to execute the script locally."


.PHONY: build macro run help

build:
	oooscript compile --pyz-out --embed --config "$(PWD)/config.json" --embed-doc "$(PWD)/odev_add_range.ods" --build-dir "build/$(BUILD_DIR)"
	rm -f $(BUILD)/*.pyz
	rm -f $(BUILD)/*.py

macro:
	oooscript compile --config "$(PWD)/config.json" --build-dir "macro"

run:
	python run.py
