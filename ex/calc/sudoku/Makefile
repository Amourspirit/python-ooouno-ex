BUILD_DIR  = sudoku
BUILD_ROOT_DIR = ../../../build
BUILD  = $(BUILD_ROOT_DIR)/$(BUILD_DIR)

help:
	@echo "Run make build to compile into 'build/$(BUILD_DIR)' folder."

.PHONY: build help

build:
	oooscript compile --embed --config "$(PWD)/config.json" --embed-doc "$(PWD)/calc-sudoku.ods" --build-dir "build/$(BUILD_DIR)"
	rm -f $(BUILD)/*.pyz
	rm -f $(BUILD)/*.py
