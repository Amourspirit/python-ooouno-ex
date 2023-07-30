help:
	@echo "Run make build to compile into 'build/cell_events' folder."

.PHONY: build help

build:
	oooscript compile --embed --config "$(PWD)/config.json" --embed-doc "$(PWD)/cell_events.ods" --build-dir "build/cell_events"
