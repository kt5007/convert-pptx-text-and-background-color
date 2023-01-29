# What is convert-pptx-text-and-background-color?
This python script convert your pptx file's text color to black and backgroud color to white.You can't convert ppt files because python-pptx doesn't support ppt files.

# You can convert file by following steps.
- Copy the files you want to convert into the taget_files direcotry
- `docker-compose up -d`
- `docker-compose exec php bash`
- `cd opt`
- `python3 main.py`
- Converted files will be made in converted_files direcotry and completed files wil be move from target_files directory to done_files direcory.
- If there have some errors in processing, error messesages will be written into exception.txt.