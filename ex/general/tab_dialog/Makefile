BUILD_DIR  = tab_dialog
BUILD_ROOT_DIR = ../../../build
BUILD  = $(BUILD_ROOT_DIR)/$(BUILD_DIR)


help:
	@echo "Run make build to compile into 'build/tab_dialog' folder."
	@echo "Run make run to execute the dialog script locally."


.PHONY: build run help

build:
# Set the current working directory
	
	cd $(ROOT_DIR) && \
	oooscript compile --pyz-out --embed --config "$(PWD)/config.json" --embed-doc "$(PWD)/tab_dialog.odt" --build-dir "build/$(BUILD_DIR)"
	rm -f $(BUILD)/*.pyz
	rm -f $(BUILD)/*.py

run:
	python $(PWD)/run.py