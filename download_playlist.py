# download_youtube_queue.py

import youtube_dl
import sys

FLAG = False


class MyLogger(object):
    def debug(self, msg):
        pass

    def warning(self, msg):
        pass

    def error(self, msg):
        print(msg)


# Main method that drives the code
def main():
    ydl_opts = {
        'format': 'bestaudio/best',
        'postprocessors': [{
            'key': 'FFmpegExtractAudio',
            'preferredcodec': 'mp3',
        }],
        'logger': MyLogger(),
        'progress_hooks': [my_hook],
        'outtmpl': 'C:/Users/magic/Music/%(title)s.%(ext)s',
    }

    with youtube_dl.YoutubeDL(ydl_opts) as ydl:

        try:
            ydl.download([sys.argv[1]])
        except:
            # ydl.download(['https://www.youtube.com/playlist?list=PLB-g8PA1MbZnhfRVEk6pLrdw4PHT4UCZ7'])
            print("You need to pass in a public Youtube Playlist URL as an arguement.")


# Method for showing progess of download/conversion
def my_hook(d):
    global FLAG

    if not FLAG and d['status'] == 'downloading':
        FLAG = not FLAG
        print(f"Downloading {d['filename'].split(chr(92))[-1][:-5]}...")
    if d['status'] == 'finished':
        FLAG = False
        print('Now converting to MP3...')


if __name__ == '__main__':
    print('Running download_youtube_queue.py')
    main()