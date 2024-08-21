# Variables
docm=cv.docm
vba_dir=src

# Default target
all: compress

# Target to extract the .docm file and VBA code using your Python script
extract: extract_vba

# Target to extract VBA code using your Python script
extract_vba:
	@mkdir -p $(vba_dir)
	@python3 python_build/extract_vba.py $(docm) $(vba_dir)

# Target to compress the extracted content back into the .docm file using your Python script
compress: insert_vba
	@python3 python_build/compile_vba.py $(vba_dir) $(docm)

# Target to reinsert VBA code into the .docm file using your Python script
insert_vba:
	@python3 python_build/compile_vba.py $(vba_dir) $(docm)

# Clean up the extraction directory
clean:
	@rm -rf $(vba_dir)
