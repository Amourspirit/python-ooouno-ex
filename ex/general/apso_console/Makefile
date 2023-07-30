help:
	@echo "Run make build to compile into 'build/apso_console' folder."
	@echo "Run make macro to compile into 'macro/apso_example.py' file."


.PHONY: build macro help

build:
	oooscript compile --embed --config "$(PWD)/config.json" --embed-doc "$(PWD)/apso_example.odt" --build-dir "build/apso_console"

macro:
	oooscript compile --config "$(PWD)/config.json" --build-dir "macro"