import re

# Testing vocab regular expression to
# catch phrase as a vocab word

# Words to Know - set up as mock "line"
line = """
audience -- plus some, more, extra words -- for a phrase – a person or group addressed by a written or spoken text
colloquialism – an informal, conversational word or phrase
conventions – rules and standards for written communication
formal language – grammar and word choice that are appropriate for a professional audience, without slang or conversational phrases
jargon – words or phrases used in specific professions, disciplines, or institutions
purpose – an author's primary reason for writing a text, most often to inform, persuade, entertain, or describe
"""
# audience – a person or group addressed by a written or spoken text
# audience, some - more words for a phrase – a person or group addressed by a written or spoken text
# audience plus some more words for a phrase – a person or group addressed by a written or spoken text

# Line: 203
# New vocab catches comma and dash if present
vocab = re.compile(r"(([\w,-]+ ){1,})–(.+)")
# Current/old
# vocab = re.compile(r"(([\w\-]+ ){1,3})–(.+)")

# Line: 295
vocab_results = {}

# Line: 304
# New "vocab_results" catches phrase or single word for vocab word
v = vocab.match(line.strip())
if v is not None:
    # New
    vocab_results[v[1]] = v[3]
    # Current/old
    # vocab_results[v[2]] = v[3]


print(f"vocab_results: {vocab_results}")
# print(f"v[0]: {v[0]}, v[1]: {v[1]}, v[2]: {v[2]}, v[3]: {v[3]}")

# # testing
# for key, value in vocab_results.items():
#     print(key + ": " + value)
#     # print(v)


# #############################################################
# def brk_if():
#     test_stop = input("\nContinue loop (y/n)? ")
#     test_stop = test_stop.lower()
#     if test_stop == 'y':
#         pass
#     else:
#         sys.exit(None)

# Words to Know - set up "line"
# line = """
# audience – a person or group addressed by a written or spoken text
# colloquialism – an informal, conversational word or phrase
# conventions – rules and standards for written communication
# formal language – grammar and word choice that are appropriate for a professional audience, without slang or conversational phrases
# jargon – words or phrases used in specific professions, disciplines, or institutions
# purpose – an author's primary reason for writing a text, most often to inform, persuade, entertain, or describe
# """


# Line: 203
# vocab = re.compile(r"(([\w\-]+ ){1,3})–(.+)")

# Line: 295
# vocab_results = {}

# Line: 304
# v = vocab.match(line.strip())
# if v is not None:
#     vocab_results[v[2]] = v[3]
