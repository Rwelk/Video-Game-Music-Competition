# custom_exceptions.py

# Raised when pp_maker.py is run without any flags.
class NoFlagsError(Exception):
	def __init__(self):
		super().__init__("\033[91mYou have to use flags to tell the program how many Rounds and Tracks you want.\nTry running \033[93mpp_maker.py -h\033[91m to learn more.\033[0m")

# Raised when pp_maker.py uses both -d and either/both of the -r and -t flags.
class AmbiguousNumRoundsTracksError(Exception):
	def __init__(self):
		super().__init__("\033[91mThe --default and (--rounds and/or --tracks) flags are mutually exclusive and cannot be used together.\033[0m")