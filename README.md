# What is convert-pptx-text-and-background-color?
This python script convert your pptx file's text color to black and backgroud color to white.You can't convert ppt files because python-pptx doesn't support ppt files.

# You can convert file by following steps.
## Make direcotries
- `mkdir opt/target_files`
- `mkdir opt/converted_files`
- `mkdir opt/done_files`
- Copy the files you want to convert into the taget_files direcotry
## Create container
- `docker-compose up -d`
## Enter into container
- `docker-compose exec python3 bash`
- #/ `cd opt`
- #/ `python3 main.py`
## Proccesing results
- Converted files will be made in converted_files direcotry and completed files wil be move from target_files directory to done_files direcory.
- If there have some errors in processing, error messesages will be written into exception.txt.