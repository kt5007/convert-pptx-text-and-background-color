import pptx
import collections.abc
import os
import with_picture_convert
import use_template_slide_convert
import shutil

if __name__ == '__main__':
    # for file_name in os.listdir('/root/opt/target_files'):
    # save_folder = 'converted_files'
    save_folder = './samples_converted'
    # target_path = '/root/opt/converted-to-pptx-files'
    target_path = './samples'
    # done_file_path = '/root/opt/pptx-done-file'
    done_file_path = './samples_done'
    for file_name in os.listdir(target_path):
        try:
            # target_path = os.path.join('./target_files/',file_name)
            target_file_path = os.path.join(target_path,file_name)
            print(target_file_path)
            has_picture = use_template_slide_convert.use_template_slide_convert(target_file_path, file_name, save_folder)
            if has_picture:
                with_picture_convert.with_picture_convert(target_file_path,'./'+save_folder + '/' + file_name)
            shutil.move(target_file_path, os.path.join(done_file_path+'/'+file_name))
        except Exception as e:
            with open('exception.txt', 'a') as f:
                print(file_name, file=f)
                print(repr(e), file=f)