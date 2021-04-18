import boto3
import configparser
import os
import datetime

config = configparser.ConfigParser()
config.read('config.cfg')

import logging
from botocore.exceptions import ClientError
from botocore.config import Config

def upload_file(file_name, bucket, object_name=None):
    """Upload a file to an S3 bucket

    :param file_name: File to upload
    :param bucket: Bucket to upload to
    :param object_name: S3 object name. If not specified then file_name is used
    :return: True if file was uploaded, else False
    """

    # If S3 object_name was not specified, use file_name
    if object_name is None:
        object_name = file_name

    # Upload the file
    s3_config = Config(
        region_name = "ap-southeast-1"
    )
    s3_client = boto3.client("s3", config=s3_config,
        aws_access_key_id = config["CREDENTIALS"]["AWS_ACCESS_KEY_ID"],
        aws_secret_access_key = config["CREDENTIALS"]["AWS_SECRET_ACCESS_KEY"])
    try:
        response = s3_client.upload_file(file_name, bucket, object_name)
    except ClientError as e:
        logging.error(e)
        return False
    return True

def main():

    for root, dirs, files in os.walk(config['PATH']['LOCAL']):
        for file in files:
            upload_file(os.path.join(root, file),
                        config['PATH']['S3_BUCKET'],
                        os.path.join( config['PATH']['S3_KEY'], root.split("\\")[-1], file).replace('\\','/')
                    )

if __name__ == "__main__":
    main()