import logging
import os
from dotenv import load_dotenv

load_dotenv()

log_file = os.environ.get("LOG_FILE", "logs.log")
logging.basicConfig(filename=log_file,
                    format='%(asctime)s %(message)s',
                    filemode='a')

logger = logging.getLogger()

logger.setLevel(logging.DEBUG)
