from dotenv import load_dotenv
import os

load_dotenv()

OPENROUTER_API_KEY = os.getenv('OPENROUTER_API_KEY')

if OPENROUTER_API_KEY is None:
    raise ValueError('OPENROUTER API KEY IS NOT FOUND')