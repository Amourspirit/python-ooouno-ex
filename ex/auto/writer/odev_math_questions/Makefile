help:
	@echo "Run make build to compile into 'build/make_formula' folder."
	@echo "Run make start to execute the start script locally"


.PHONY: build help, msg-short, msg-long, msg-warn, msg-error, msg-yes

build:
	oooscript compile --embed --config "$(PWD)/config.json" --embed-doc "$(PWD)/data/blank.odt" --build-dir "build/make_formula"

start:
	python $(PWD)/start.py
