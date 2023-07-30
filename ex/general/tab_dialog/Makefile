ROOT_DIR = ../../../


help:
	@echo "Run make build to compile into 'build/tab_dialog' folder."
	@echo "Run make run to execute the dialog script locally."


.PHONY: build run help

build:
# Set the current working directory
	
	cd $(ROOT_DIR) && \
	oooscript compile --embed --config "$(PWD)/config.json" --embed-doc "$(PWD)/tab_dialog.odt" --build-dir "build/tab_dialog"

run:
	python $(PWD)/run.py