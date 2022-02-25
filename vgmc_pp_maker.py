# vgmc_pp_maker.py

from pptx import Presentation
from pptx.util import Inches

import argparse
import sys
from pathlib import Path
ROOT = Path(__file__).parent.absolute()

from datetime import date

# Parse the arguments provided to argparse.
parser = argparse.ArgumentParser(description="Use this file to create a PowerPoint for the Florida Southern College Computer Science Club's Video Game Music Competiton (FSCCSCVGMC).")
parser.add_argument('-d', '--default', action="store_true", help='creates a default PowerPoint with 5 Rounds each with 10 Tracks.')
parser.add_argument('-r', '--rounds', type=int, nargs='?', help='provide a number of rounds')
parser.add_argument('-t', '--tracks', type=int, nargs='?', help='provide a number of tracks for each round')
args = parser.parse_args()

def main():
	
	# Determine how many Rounds and Tracks per Round there will be.
	rounds, tracks = determine_parameters()
	print(f"There will be {rounds} Rounds Each with {tracks} Tracks.")

	# Create the PowerPoint.
	prs = Presentation()
	prs.slide_width = Inches(16)
	prs.slide_height = Inches(9)

	# Create the different Slide Layouts
	title_slide = prs.slide_layouts[0]
	title_content_slide = prs.slide_layouts[1]
	section_header_slide = prs.slide_layouts[2]


	prs = title_and_rules(prs, title_slide, title_content_slide, rounds, tracks)





	# Save the Presentation
	save_presentation(prs)


# This method reads the arguements passed in to determine how many Rounds and Tracks there will be.
def determine_parameters():

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

	return rounds, tracks


# This method creates the Title and Rules Slides.
def title_and_rules(prs, title_slide_layout, rules_slide_layout, rounds, tracks):

	# Create the Title Slide.
	slide1 = prs.slides.add_slide(title_slide_layout)

	# Populate the slide's Title and Subtitle.
	set_title(slide1, "Video Game Music Guessing Competition")
	set_text(slide1, "Sponsored by the Computer Science Club")

	# Create the Rules Slide.
	slide2 = prs.slides.add_slide(rules_slide_layout)

	# Populate the slide's Title and Content.
	set_title(slide2, "Rules and Scoring")
	set_text(slide2, f"There will be {rounds} Rounds of {tracks} Tracks")
	
	add_text(slide2, "You will get roughly 30 – 45 seconds of music to guess from.")
	add_text(slide2, "Scoring is as follows:")
	add_text(slide2, "1 point for Game Franchise", 1)
	add_text(slide2, "1 point for Specific Game", 1)
	add_text(slide2, "1 point for Track Name/Place", 1)
	add_text(slide2, "A couple songs don’t have official releases, or play in multiple places. Those have a star listed on the answer key, and if you put something close to it you’ll still receive the point.")

	return prs


# This method is for more easily setting the title of a slide.
def set_title(slide, text):
	slide.shapes.title.text = text


# This method is for more easily setting the first line of text in a slide.
def set_text(slide, text):
	slide.placeholders[1].text = text


# This method is for adding additional lines of text to a slide
def add_text(slide, text, level=0):

	p = slide.shapes.placeholders[1].text_frame.add_paragraph()
	p.text = text

	if level: p.level = level



# This method is for saving the PowerPoint produced by this script.
def save_presentation(presentation):
	
	# Get the current year and season
	today = date.today()
	year = today.year
	season = "Spring" if today.month < 7 else "Fall"

	# Save the passed in presentation.
	presentation.save(ROOT / f"VGMC {year} {season} Slides.pptx")


if __name__ == '__main__':
	print('Running vgmc_pp_maker.py')
	main()