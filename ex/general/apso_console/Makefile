BUILD_DIR  = apso_console
BUILD_ROOT_DIR = ../../../build
BUILD  = $(BUILD_ROOT_DIR)/$(BUILD_DIR)

help:
	@echo "Run make build to compile into 'build/$(BUILD_DIR)' folder."
	@echo "Run make macro to compile into 'macro/apso_example.py' file."


.PHONY: build macro help

build:
	oooscript compile --embed --config "$(PWD)/config.json" --embed-doc "$(PWD)/apso_example.odt" --build-dir "build/$(BUILD_DIR)"
	rm -f $(BUILD)/*.py

macro:
	oooscript compile --config "$(PWD)/config.json" --build-dir "macro"