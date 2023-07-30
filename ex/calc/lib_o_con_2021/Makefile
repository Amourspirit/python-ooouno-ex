MYDIR = res
BUILD_DIR = ../../../build/lib_o_con_2021

help:
	@echo "Run make build to compile into 'build/lib_o_con_2021' folder."

.PHONY: build help

files:
	rsync -rupE $(MYDIR) $(BUILD_DIR)/

build: files
	oooscript compile --embed --config "$(PWD)/config.json" --embed-doc "$(PWD)/lib_o_con_2021.ods" --build-dir "build/lib_o_con_2021"