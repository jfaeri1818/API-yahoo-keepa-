

print("Random ASCII Art:")

def random_ascii_art(rows, cols):
  for _ in range(rows):
    line = ''.join(random.choice([' ', '*']) for _ in range(cols))
    print(line)

random_ascii_art(10, 40)



