help:
	@echo "Run make build to compile into './data/dialog_menu.ods' file."
	@echo "Run make start to execute the dialog script locally"


.PHONY: build help,

build:
	oooscript compile --embed --config "$(PWD)/config.json" --embed-doc "$(PWD)/data/src_doc/src_doc.ods" --build-dir "$(PWD)/data"
	rm -f "$(PWD)/data/dialog_menu.py"

start:
	python $(PWD)/start.py
