# Define the file name
file_name = 'dialogue.txt'

# Initialize a dictionary to hold character counts for each character
char_count = {'Carlson': 0, 'Dylan': 0, 'Denzel': 0}
total_count = 0  # Initialize total count of Chinese characters

# Read the dialogue from the text file
with open(file_name, 'r', encoding='utf-8') as file:
    dialogue = file.readlines()

# Process each line in the dialogue
for line in dialogue:
    # Strip any leading/trailing whitespace and check for character name
    line = line.strip()
    if line.endswith(':'):
        current_character = line[:-1]  # Get the character name (removing ':')
    elif current_character in char_count:
        # Count Chinese characters in the line
        for char in line:
            if '\u4e00' <= char <= '\u9fff':  # Range for common Chinese characters
                char_count[current_character] += 1
                total_count += 1  # Increment total count

# Print the character count for each character
for character, count in char_count.items():
    print(f"{character}: {count} Chinese characters")

# Print the total count of Chinese characters
print(f"Total Chinese characters: {total_count}")
