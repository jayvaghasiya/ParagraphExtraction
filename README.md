# Extract Text Blocks from PDF
Extract text blocks from a PDF document with each paragraph as a separate line and text being in reading order going from first column from top to bottom and then into second column and then third column. Dump the output in an excel file.

## Prerequisites

- Python 3.10+
- PyMuPDF
- OpenPyXL

## Usage

$ python extract_text_blocks.py --input input_file.pdf -o output_file.xlsx


### Arguments

- `input_file`: Path to the input PDF file. This is a required argument.
- `output_file`: Path to the output Excel file. This is an optional argument. Default value is `output.xlsx`.

## Example

To extract text blocks from a PDF file named `keppel-corporation-limited-annual-report-2018.pdf` and save the output to an Excel file named `sample.xlsx`, run the following command:

$ python extract_text_blocks.py --input input/keppel-corporation-limited-annual-report-2018.pdf -o sample.xlsx


The output Excel file will be saved in the current working directory.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
