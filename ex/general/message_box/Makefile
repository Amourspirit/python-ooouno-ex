BUILD_DIR  = message_box
BUILD_ROOT_DIR = ../../../build
BUILD  = $(BUILD_ROOT_DIR)/$(BUILD_DIR)


help:
	@echo "Run make build to compile into 'build/$(BUILD_DIR)' folder."
	@echo "Run make msg-short to execute the dialog script locally with short message."
	@echo "Run make msg-long to execute the dialog script locally with long message."
	@echo "Run make msg-warn to execute the dialog script locally with warning message."
	@echo "Run make msg-error to execute the dialog script locally with error message."
	@echo "Run make msg-yes to execute the dialog script locally with yes/no message."


.PHONY: build help, msg-short, msg-long, msg-warn, msg-error, msg-yes

build:
	oooscript compile  --pyz-out --embed --config "$(PWD)/config.json" --embed-doc "$(PWD)/msgbox.odt" --build-dir "build/$(BUILD_DIR)"
	rm -f $(BUILD)/*.pyz
	rm -f $(BUILD)/*.py

msg-short:
	python $(PWD)/run.py "short"

msg-long:
	python $(PWD)/run.py "long"

msg-warn:
	python $(PWD)/run.py "warn"

msg-error:
	python $(PWD)/run.py "error"

msg-yes:
	python $(PWD)/run.py