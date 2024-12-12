BUILD_DIR  = dialog_menu
BUILD_ROOT_DIR = ../../../build
BUILD  = $(BUILD_ROOT_DIR)/$(BUILD_DIR)

help:
	@echo "Run make build to compile into 'build/$(BUILD_DIR)' file."
	@echo "Run make start to execute the dialog script locally"


.PHONY: build help,

build:
	oooscript compile --pyz-out --embed --config "$(PWD)/config.json" --embed-doc "$(PWD)/data/src_doc/src_doc.ods" --build-dir "build/$(BUILD_DIR)"
	rm -f $(BUILD)/*.pyz
	rm -f $(BUILD)/*.py

start:
	python $(PWD)/start.py
