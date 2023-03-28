include .env

timetable-to-json:
	cd timetable && \
	python scripts/main.py --import_path $(IMPORT_PATH) --export_path $(EXPORT_PATH) --min_col $(MIN_COL) --max_col $(MAX_COL) --min_row $(MIN_ROW)
syllabus-to-json:
	cd syllabus && \
	python scripts/main.py --import_path $(IMPORT_PATH) --export_path $(EXPORT_PATH) --config_path $(CONFIG_PATH)