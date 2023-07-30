help:
	@echo "Run make build to compile into 'build/add_range_data' folder."
	@echo "Run make macro to compile into 'macro/odev_add_range.py' file."


.PHONY: build macro help

build:
	oooscript compile --embed --config "$(PWD)/config.json" --embed-doc "$(PWD)/odev_add_range.ods" --build-dir "build/add_range_data"

macro:
	oooscript compile --config "$(PWD)/config.json" --build-dir "macro"