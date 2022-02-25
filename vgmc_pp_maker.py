# vgmc_pp_maker.py

from pptx import Presentation
from pptx.util import Inches, Pt

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
	prs.slide_width = Inches(13.333)
	prs.slide_height = Inches(7.5)

	# Create the different Slide Layouts
	title_slide = prs.slide_layouts[0]
	title_content_slide = prs.slide_layouts[1]
	section_header_slide = prs.slide_layouts[2]

	# Create the Title and Rules Slides
	title_and_rules(prs, title_slide, title_content_slide, rounds, tracks)





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
	add_text(slide1, "Sponsored by the Computer Science Club", size=18)

	# Create the Rules Slide.
	slide2 = prs.slides.add_slide(rules_slide_layout)

	# Populate the slide's Title and Content.
	set_title(slide2, "Rules and Scoring")
	add_text(slide2, f"There will be {rounds} Rounds of {tracks} Tracks", size=20)
	
	add_text(slide2, "You will get roughly 30 – 45 seconds of music to guess from.", size=20)
	add_text(slide2, "Scoring is as follows:", size=20)
	add_text(slide2, "1 point for Game Franchise", level=1)
	add_text(slide2, "1 point for Specific Game", level=1, italic=True)
	add_text(slide2, "1 point for Track Name/Place", level=1, bold=True)
	add_text(slide2, "A couple songs don’t have official releases, or play in multiple places. Those have a star listed on the answer key, and if you put something close to it you’ll still receive the point.", size=20)


# This method is for more easily setting the title of a slide.
def set_title(slide, text, size=54):
	t = slide.shapes.title.text_frame.paragraphs[0]
	
	t.font.size = Pt(size)
	t.text = text


# This method is for adding additional lines of text to a slide
def add_text(slide, text, level=0, size=18, bold=False, italic=False):

	# Each slide has some number of items on them called shapes, which themselves store some number of
	#     placeholders.
	# These placeholders are stored in a dictionary simply called placeholders{}.
	# To add text to a slide, we have to access the placeholder stored at key 1.
	# Note that since placeholders is not an array, this is not accessing INDEX 1 but KEY 1.
	# From there, we can acess the text_frame, which is what stores the text in the shape.
	content_area = slide.shapes.placeholders[1].text_frame

	# text_frames are composed of various Paragraphs.
	# p will be the Paragraph that our text argument will be written into.
	# However, before we can write we have to determine whether the text_frame already has content.
	# If it does, we have to call .add_paragraph(), a method belonging to text_frame that creates a new
	#     Paragraph that can be written to.
	# Otherwise, we should write to the Paragraph already there.
	p = content_area.add_paragraph() if content_area.text else content_area.paragraphs[0]

	# Apply text styling.
	p.level = level
	p.font.size = Pt(size)
	p.font.bold = bold
	p.font.italic = italic

	# Write the text into the space.
	p.text = text


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