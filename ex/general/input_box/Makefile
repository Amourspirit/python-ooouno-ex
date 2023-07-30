help:
	@echo "Run make build to compile into 'build/input_box' folder."
	@echo "Run make run to execute the dialog script locally."


.PHONY: build run help

build:
	oooscript compile --embed --config "$(PWD)/config.json" --embed-doc "$(PWD)/inputbox.odt" --build-dir "build/input_box"

run:
	python $(PWD)/run.py