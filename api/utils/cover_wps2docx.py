import subprocess
import logging
import os
import platform
import time
import tempfile


# 配置日志
logger = logging.getLogger(__name__)
#########安装libreoffice#######################
#sudo apt-get update
#sudo apt install libreoffice
##验证是否安装成功
##libreoffice --version

####如果报如下错误，需要安装dialog组件：debconf: unable to initialize frontend: Dialog
###sudo apt-get install dialog

##soffice --headless --convert-to docx W020240423576340405124.wps

def convert_to_docx(input_file_path, output_file_path):
    with open(input_file_path, 'rb') as f:
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as temp_docx:
            command = f"libreoffice --headless --convert-to docx:writer_docx_Export --outdir {os.path.dirname(output_file_path)} {input_file_path}"
            subprocess.run(command, shell=True)
            # 移动临时文件到最终位置
            os.rename(temp_docx.name, output_file_path)

def execute_libreoffice_command(command):
    try:
        logger.debug(f"execute_libreoffice_command cmd : {command}")
        process = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        # 等待命令执行完成
        stdout, stderr = process.communicate()

        # 检查进程退出状态
        exit_status = process.returncode
        logger.debug(f"i----{exit_status}")

        if exit_status != 0:
            logger.error(f"execute_libreoffice_command cmd exitStatus {exit_status}")
            if stderr:
                logger.error(f"Error: {stderr.decode('utf-8')}")
            return False
        else:
            logger.debug(f"execute_libreoffice_command cmd exitStatus {exit_status}")
            if stdout:
                logger.debug(f"Output: {stdout.decode('utf-8')}")

    except subprocess.CalledProcessError as e:
        logger.error(f"execute_libreoffice_command {command} error {e}")
        return False
    except Exception as e:
        logger.error(f"An error occurred: {e}")
        return False

    logger.info("转化结束.......")
    return True


def convert_office_to_docxorpdf(input_file, output_file,cover_file_type='docx'):

    # if cover_file_type !='docx' and cover_file_type !='pdf'
    #     return

    start_time = time.time()
    command = ""
    os_name = platform.system()
    if os_name == "Windows":
        command = f"cmd /c start soffice --headless --invisible --convert-to {cover_file_type} {input_file} --outdir {output_file}"
    else:
        command = f"libreoffice --headless --invisible --convert-to {cover_file_type} {input_file} --outdir {output_file}"

    flag = execute_libreoffice_command(command)
    end_time = time.time()
    logger.debug(f"用时: {end_time - start_time} 秒")

    return flag


def convert_stream_to_format(input_stream, output_format):
    # 创建一个临时文件
    with tempfile.NamedTemporaryFile(delete=False) as tmp_file:
        # 将文件流写入临时文件
        tmp_file.write(input_stream)
        tmp_file.flush()

        try:
            # 构建 LibreOffice 命令
            command = f"libreoffice --headless --convert-to {output_format} --outdir {os.path.dirname(tmp_file.name)} {tmp_file.name}"
            subprocess.run(command, shell=True, check=True)

            # 构建输出文件的路径
            output_file = os.path.splitext(tmp_file.name)[0] + f".{output_format.split(':')[-1]}"

            # 返回转换后的文件内容
            with open(output_file, 'rb') as f:
                return f.read()
        except Exception as e:
            logger.error(f"convert_stream_to_format error occurred: {e}")
            return None

        finally:
            # 删除临时文件
            os.unlink(tmp_file.name)
            if os.path.exists(output_file):
                os.unlink(output_file)


# if __name__ == "__main__":

#     # 使用示例
#     # input_file_path = 'W020240722613678243393.wps'
#     # out_file_path= 'W020240722613678243393.docx'
#     # output_format = "docx"  # 或者 "docx"
    
#     input_file_path = '20141117165909284.xls'
#     out_file_path= '20141117165909284.xlsx'
#     output_format = "xlsx"  # 或者 "docx"
    
            
#     with open(input_file_path, 'rb') as f:
#         input_stream = f.read()
#         output_stream = converted_content = convert_stream_to_format(input_stream, output_format)
#         with open(out_file_path, 'wb') as tmp_file:
#             # 将文件流写入临时文件
#             tmp_file.write(output_stream)
#             tmp_file.flush()
        
    