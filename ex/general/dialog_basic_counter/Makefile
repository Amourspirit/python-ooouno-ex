BUILD_DIR  = BasicCounter
BUILD_ROOT_DIR = ../../../build
BUILD  = $(BUILD_ROOT_DIR)/$(BUILD_DIR)

help:
	@echo "Run make build to compile into 'build/BasicCounter' folder."


.PHONY: build help

build:
	oooscript compile --embed --config "$(PWD)/config.json" --embed-doc "$(PWD)/BasicCounter.ods" --build-dir "build/$(BUILD)"
	rm -f $(BUILD)/*.py
