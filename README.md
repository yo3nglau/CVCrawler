This program is developed to crawl **CVPR**, **ICCV**, **NeurIPS**, **ICLR**, **ECCV**, **ICML** paper abstracts into docx/pdf files using [Scrapy](https://github.com/scrapy/scrapy), [Requests](https://github.com/psf/requests), and [python-docx](https://github.com/python-openxml/python-docx).

## Supported conferences and corresponding years

|         |        2024        |        2023        |           2022           |           2021           |
| :-----: | :----------------: | :----------------: | :----------------------: | :----------------------: |
|  CVPR   |      upcoming      | :heavy_check_mark: |    :heavy_check_mark:    |    :heavy_check_mark:    |
|  ICCV   | :heavy_minus_sign: | :heavy_check_mark: |    :heavy_minus_sign:    |    :heavy_check_mark:    |
| NeurIPS |      upcoming      | :heavy_check_mark: |    :heavy_check_mark:    |    :heavy_check_mark:    |
|  ICLR   | :heavy_check_mark: | :heavy_check_mark: |    :heavy_check_mark:    |    :heavy_check_mark:    |
|  ICML   |      upcoming      | :heavy_check_mark: | :heavy_multiplication_x: | :heavy_multiplication_x: |
|  ECCV   |      upcoming      | :heavy_minus_sign: |    :heavy_check_mark:    |    :heavy_minus_sign:    |

:heavy_check_mark: denotes that it is available.

:heavy_minus_sign: denotes that it is missing because the conference is held biennially.

:heavy_multiplication_x: denotes that it has not been supported by the program thus far.

## Environment

Create a conda virtual environment and activate it:

```python
conda create -n cvcrawler python=3.10 -y
conda activate cvcrawler
```

Install requirements:

```python
pip install -r requirements/requirements.txt
```

If toc and pdf are required (ONLY support Windows system installed MS Word), install optional requirements:

```python
pip install -r requirements/optional.txt
```

## Crawl data into json logits (optional)

The json logits have been stored in advance, run these commands to crawl from scratch if interest.

If previous years are required, ensure that change `logits json filename` and `year` **simultaneously**.

### CVPR, ICCV, ECCV

```python
scrapy crawl cvf_paper_spider -O logits/CVPR_2023_Abstract.json -a conference=CVPR -a year=2023
```

```python
scrapy crawl cvf_paper_spider -O logits/ICCV_2023_Abstract.json -a conference=ICCV -a year=2023
```

```python
scrapy crawl eccv_paper_spider -O logits/ECCV_2022_Abstract.json -a year=2022
```

### NeurIPS, ICLR, ICML

For these conferences, the program will automatically crawl json logits based on OpenReview URLs if they are not locally available.

## Parse data into docx/pdf files (off-the-shelf)

```python
usage: output.py [-h] [--conference CONFERENCE] [--year YEAR] [--toc] [--pdf] [--wps] [--select] [--all]
                                                                                                        
parse CVF or OpenReview paper abstracts to a docx file                                                  
                                                                                                        
options:                                                                                                
  -h, --help            show this help message and exit                                                 
  --conference CONFERENCE
                        exact abbreviation of conference, support NeurIPS, ICLR, ICML, CVPR, and ICCV so far
  --year YEAR           year of conference
  --toc                 add table of contents at the very beginning, ONLY support Windows system installed MS Word, because of the `win32com` package
  --pdf                 additionally generate pdf, ONLY support Windows system installed MS Word, because of the `win32com` package
  --wps                 use WPS to convert pdf, which performs better than MS Word and can handle large file size
  --select              interactively select specific accepted types, for OpenReview papers
  --all                 lazy export paper abstracts of all conferences in all supported years
```

Select commands below that suit you most.

### Plain results

Generate all paper abstracts into docx file:

```python
python output.py --conference CONFERENCE --year YEAR
```

### Results with TOC

Generate all paper abstracts into docx file, and display table of contents at the very beginning:

```python
python output.py --conference CONFERENCE --year YEAR --toc
```

**Note:** ONLY support Windows system installed MS Word, because of the `win32com` package.

### Selected results

Generate interactively selected paper abstracts into docx file (suitable for NeurIPS, ICLR, ICML):

```python
python output.py --conference CONFERENCE --year YEAR --toc --select
```

### PDF results

Generate all paper abstracts into pdf file:

```python
python output.py --conference CONFERENCE --year YEAR --toc --select --pdf
```

**Note:** ONLY support Windows system installed MS Word, because of the `win32com` package.

:fire: **If WPS is available:**

```python
python output.py --conference CONFERENCE --year YEAR --toc --select --pdf --wps
```

> Practically, WPS performs better than MS Word when convert docx file into pdf file. Specifically, WPS-generated hyperlinks to the table of contents allow to pinpoint the exact location of the papers, while MS Word-generated ones only locate the page of the papers. Besides, WPS succeeds to handle large docx file while MS Word not.

### Lazy export

Generate paper abstracts of all conferences in all supported years.

```python
python output.py --toc --pdf --wps --all
```

