# vgmc_pp_maker.py

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_VERTICAL_ANCHOR
from math import ceil


import copy
from pptx.shapes.autoshape import Shape

import argparse
from io import FileIO
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
	num_rounds, num_tracks = determine_parameters()
	print(f"\033[96mThere will be \033[93m{num_rounds}\033[96m Rounds Each with \033[93m{num_tracks}\033[96m Tracks.\033[0m")

	# Create the PowerPoint using the slide template.
	prs = Presentation(ROOT / "template.pptx")

	# Create the Title and Rules Slides
	rules_slide(prs, num_rounds, num_tracks)

	for round_num in range(1, num_rounds + 1):

		round_slide(prs, round_num)

		for track_num in range(1, num_tracks + 1):
			track_slide(prs, round_num, track_num)

		review_slide(prs, round_num, num_tracks)



	# Save the Presentation
	save_presentation(prs)


# This method reads the arguements passed in to determine how many Rounds and Tracks there will be.
def determine_parameters():

	# default is supposed to be mutually exclusive with rounds and tracks, so if -d is provided, it cannot
	#	 be combined with -r or -t.
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


# This method is for generating a new slide.
def add_slide(prs, layout, title):
	
	# Create the slide based off of the passed in layout.
	slide = prs.slides.add_slide(layout)
	
	# Title the slide.
	slide.shapes.title.text_frame.paragraphs[0].text = title

	return slide


# This method is for adding additional lines of text to a slide
def add_text(slide, text, level=0, bold=False, italic=False):

	# Each slide has some number of items on them called shapes, which themselves store some number of
	#	 placeholders.
	# These placeholders are stored in a dictionary simply called placeholders{}.
	# To add text to a slide, we have to access the placeholder stored at key 1.
	# Note that since placeholders is not an array, this is not accessing INDEX 1 but KEY 1.
	# From there, we can acess the text_frame, which is what stores the text in the shape.
	content_area = slide.shapes.placeholders[1].text_frame

	# text_frames are composed of various Paragraphs.
	# p will be the Paragraph that our text argument will be written into.
	# However, before we can write we have to determine whether the text_frame already has content.
	# If it does, we have to call .add_paragraph(), a method belonging to text_frame that creates a new
	#	 Paragraph that can be written to.
	# Otherwise, we should write to the Paragraph already there.
	p = content_area.add_paragraph() if content_area.text else content_area.paragraphs[0]

	# Apply text styling.
	p.level = level
	p.font.bold = bold
	p.font.italic = italic

	# Write the text into the space.
	p.text = text


# This method creates the Rules Slide.
def rules_slide(prs, rounds, tracks):

	# Create the Rules Slide.
	slide = add_slide(prs, prs.slide_layouts[1], "Rules and Scoring")

	# Add text to the Rules Slide.
	add_text(slide, f"There will be {rounds} Rounds of {tracks} Tracks")
	add_text(slide, "You will get roughly 30 – 45 seconds of music to guess from.")
	add_text(slide, "Scoring is as follows:")
	add_text(slide, "1 point for Game Franchise", level=1)
	add_text(slide, "1 point for Specific Game", level=1)
	add_text(slide, "1 point for Track Name/Place", level=1)
	add_text(slide, "A couple songs don’t have official releases, or play in multiple places. Those have a star listed on the answer key, and if you put something close to it you’ll still receive the point.")


# This method is for making the slides that show the round number before a round.
def round_slide(prs, round_num):

	# Create the slide that announces the start of a new Round.
	slide = add_slide(prs, prs.slide_layouts[2], f"Round {round_num}")


# This method is for making the individual track slides for each round.
def track_slide(prs, round_num, track_num):

	# Create and title the slide.
	slide = add_slide(prs, prs.slide_layouts[3], f"Track {track_num}")

	# Add the Audio object.
	slide.shapes.add_movie(
		FileIO(ROOT / f"tracks/{round_num}-{track_num}.mp3", "rb"),
		left=Inches(8), top=Inches(1.5), width=Inches(6), height=Inches(6),
		poster_frame_image=FileIO(ROOT / "speaker.png", "rb"),
		mime_type='audio/mp3',
	)




def clone_shape(shape):
	"""Add a duplicate of `shape` to the slide on which it appears."""
	shape_obj = shape.element
	sp_tree = shape_obj.getparent()
	new_sp = copy.deepcopy(shape_obj)
	sp_tree.append(new_sp)
	new_shape = Shape(new_sp, None)
	new_shape.left = shape.left
	new_shape.top = shape.top
	new_shape.width = shape.width
	new_shape.height = shape.height
	return new_shape






# This method is for making the Review slide at the end of each round.
def review_slide(prs, round_num, num_tracks):

	slide = add_slide(prs, prs.slide_layouts[4], f"Round {round_num} Review")


	audio_base_x = 5.4
	audio_base_y = 1.11

	paragraph_base_x = 7.56
	paragraph_base_y = 4

	template_text_box = slide.shapes.placeholders[10]




	
	# # Set the parameters of the text box already there.
	# first_item = slide.shapes.placeholders[1]
	# first_item.left = Inches(paragraph_base_x)
	# first_item.top = Inches(paragraph_base_y)

	# fi_tb = first_item.text_frame.paragraphs[0]
	# fi_tb.text = " Track 1"

	# Set parameters for every other item.
	for i in range(0, num_tracks):
		copied_text_box = clone_shape(template_text_box)
		copied_text_box.left = Inches(
			paragraph_base_x 
			if i % 2 == 0 else 
			paragraph_base_x + 4.49
		)

		copied_text_box.top = Inches(
			4 - (
				ceil(num_tracks / 4) - (i // 2) - (
					1 if ceil(num_tracks / 2) % 2 != 0 else 0.5
				)
			)
		)

		copied_text_box.text_frame.paragraphs[0].text = f" Track {i + 1}"


	# Finally, delete the extra template text box that was copied from.
	sp = template_text_box._sp
	sp.getparent().remove(sp)


	# for i in range(1, num_tracks):

	# 	slide.shapes.add_movie(
	# 		FileIO(ROOT / f"tracks/{round_num}-{i + 1}.mp3", "rb"),
	# 		left=Inches(audio_base_x), top=Inches((1.1 * i) + audio_base_y), width=Inches(0.9), height=Inches(0.9),
	# 		poster_frame_image=FileIO(ROOT / "speaker.png", "rb"),
	# 		mime_type='audio/mp3'
	# 	)

	# 	tb = slide.shapes.add_textbox(
	# 		left=Inches(paragraph_base_x), top=Inches((1.09 * i) + paragraph_base_y), 
	# 		width=Inches(2.68), height=Inches(0.71))
		
	# 	tb_text = tb.text_frame.paragraphs[0]
	# 	tb_text.font.size = Pt(36)
	# 	tb_text.text = f"- Track {i + 1}"


# This method is for saving the PowerPoint produced by this script.
def save_presentation(presentation):
	
	# Get the current year and season
	today = date.today()
	year = today.year
	season = "Spring" if today.month < 7 else "Fall"

	# Save the passed in presentation.
	presentation.save(ROOT / f"VGMC {year} {season} Slides.pptx")


if __name__ == '__main__':
	print('\n\033[92mRunning vgmc_pp_maker.py\033[0m')

	main()
	# try:
	# 	main()

	# except Exception as e:
	# 	print(f'\033[91m{e}\033[0m')