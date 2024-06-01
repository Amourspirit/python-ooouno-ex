help:
	@echo "Run make data to create './data/src_doc/src_doc.ods' file."
	@echo "Run make build to compile into './data/custom_props.ods' file."
	@echo "Run make start to execute the dialog script locally"


.PHONY: build help, start, data

data:
	rm -f "$(PWD)/data/src_doc/src_props.py"
	python $(PWD)/generate_data.py

build:
	oooscript compile --embed --config "$(PWD)/config.json" --embed-doc "$(PWD)/data/src_doc/src_doc.ods" --build-dir "$(PWD)/data"
	rm -f "$(PWD)/data/custom_props.py"

start:
	python $(PWD)/start.py
