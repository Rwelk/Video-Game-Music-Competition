# download_youtube_queue.py

import youtube_dl
import argparse
from pathlib import Path
ROOT = Path(__file__).parent.absolute()


# argparse code for passing in the playlist to download with -p or --playlist
parser = argparse.ArgumentParser(description="Use this file to download and convert the videos in a Youtube playlist to .mp3 files.")
parser.add_argument('-p', '--playlist', required=True, help='the link to the Youtube playlist that will be downloaded')
args = parser.parse_args()

# FLAG is a boolean flag used for the hooks that print the code's current status.
FLAG = False


# I'll be honest, I have zero clue what the MyLogger class is for.
class MyLogger(object):
    def debug(self, msg):
        pass

    def warning(self, msg):
        pass

    def error(self, msg):
        print(msg)


# Main method that drives the code
def main():
    ydl = youtube_dl.YoutubeDL(
    {
        'format': 'bestaudio/best',
        'postprocessors': [{
            'key': 'FFmpegExtractAudio',
            'preferredcodec': 'mp3',
        }],
        'logger': MyLogger(),
        'progress_hooks': [my_hook],
        'outtmpl': str(ROOT / '%(title)s.%(ext)s'),
    })

    botched_vids = []
    with ydl:
        result = ydl.extract_info(args.playlist, download=False)

        for video in result['entries']:
            try:
                ydl.download([video['webpage_url']])
            except KeyboardInterrupt:
                print("You triggered a KeyBoardInterrupt exception.")
                print("Shutting down download_yputube_queue.py")
                return
            except:
                print('Something went wrong with this song, skipping for now...')
                botched_vids.append(video['title'])

    if botched_vids:
        print("The following songs couldn't be downloaded:")

        for title in botched_vids:
            print(f' {title}')


# Method for showing progess of download/conversion
def my_hook(d):
    global FLAG

    if not FLAG and d['status'] == 'downloading':
        FLAG = not FLAG
        print(f"\nDownloading {d['filename'].split(chr(92))[-1].split('.')[0]}...")
    if d['status'] == 'finished':
        FLAG = False
        print('Now converting to MP3...')


if __name__ == '__main__':
    print('\nRunning download_youtube_queue.py')
    main()