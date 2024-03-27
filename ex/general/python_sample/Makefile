help:
	@echo "Run make build to compile into 'build/python_sample' folder."


.PHONY: build help, msg-short, msg-long, msg-warn, msg-error, msg-yes

build:
	oooscript compile --embed --config "$(PWD)/config.json" --embed-doc "$(PWD)/data/sample.odt" --build-dir "build/python_sample"
