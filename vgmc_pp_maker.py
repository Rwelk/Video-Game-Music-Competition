# vgmc_pp_maker.py

import argparse
import sys
from pathlib import Path
ROOT = Path(__file__).parent.absolute()

# Parse the arguments provided to argparse.
parser = argparse.ArgumentParser(description="Use this file to create a PowerPoint for the Florida Southern College Computer Science Club's Video Game Music Competiton (FSCCSCVGMC).")
parser.add_argument('-d', '--default', action="store_true", help='creates a default PowerPoint with 5 Rounds each with 10 Tracks.')
parser.add_argument('-r', '--rounds', type=int, nargs='?', help='provide a number of rounds')
parser.add_argument('-t', '--tracks', type=int, nargs='?', help='provide a number of tracks for each round')
args = parser.parse_args()

def main():

	# default is supposed to be mutually exclusive with rounds and tracks, so if -d is provided, it cannot
	#     be combined with -r or -t.
	# One the other end, -r and -t can be used together or separatly, but cannot be used with -d.
	if args.default and (args.rounds or args.tracks):
		print("--default and --rounds|--tracks are mutually exclusive and cannot be used together.")
		sys.exit(2)
	else:

		# If -d was provided, set rounds and tracks to the default values of 5 and 10 respectively.
		if args.default: rounds, tracks = 5, 10

		# Otherwise, check to see if -r and/or -t was/were provided.
		# If -r was provided, set rounds to that number, else set it to 5.
		# If -t was provided, set tracks to that number, else set it to 10.
		else: 
			rounds = args.rounds if args.rounds is not None else 5
			tracks = args.tracks if args.tracks is not None else 10

	print(f"There will be {rounds} Rounds Each with {tracks} Tracks.")


if __name__ == '__main__':
	print('Running vgmc_pp_maker.py')
	main()