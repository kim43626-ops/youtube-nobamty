"""Read a YouTube API key without printing or persisting its value."""
import os
VARIABLE='YOUTUBE_API_KEY'

def youtube_api_key():
    return os.environ.get(VARIABLE,'').strip()
