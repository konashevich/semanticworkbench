from dotenv import load_dotenv
from . import config
import logging
import sys

# Load environment variables from .env into the settings object.
load_dotenv()
settings = config.Settings()

# Configure logging for this package to write INFO+ to stdout instead of stderr.
# This prevents hosts that surface stderr as 'warnings' from showing normal INFO logs.
try:
	level_name = (settings.log_level or "INFO").upper()
	level = logging.getLevelNamesMapping().get(level_name, logging.INFO)
except Exception:
	level = logging.INFO

handler = logging.StreamHandler(stream=sys.stdout)
handler.setFormatter(logging.Formatter("[%(asctime)s] %(levelname)-8s %(message)s"))
root = logging.getLogger()
# Only add our handler if no handlers are present to avoid duplicate logs when running under other frameworks
if not root.handlers:
	root.addHandler(handler)
root.setLevel(level)
