class NoFlagsError(Exception):
	def __init__(self):
		super().__init__("You have to use flags to tell the program how many Rounds and Tracks you want.\nTry using the -h flag to learn more.")

class AmbiguousNumRoundsTracksError(Exception):
	def __init__(self):
		super().__init__("--default and --rounds|--tracks are mutually exclusive and cannot be used together.")