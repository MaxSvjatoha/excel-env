
'''
from fuzzywuzzy import fuzz

def find_closest_match(query, choices):
    best_match = None
    highest_score = 0
    for choice in choices:
        score = fuzz.ratio(query, choice)
        print(f"{query} vs {choice}: {score}")
        if score > highest_score:
            highest_score = score
            best_match = choice
    return best_match

choices = ['apple', 'banana', 'cherry', 'durian']
query = 'aple'

closest_match = find_closest_match(query, choices)
print(f"The closest match to '{query}' is '{closest_match}'.")
'''

a = None
a = int(a) + 1
print(a)