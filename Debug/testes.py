import hashlib
import re

def simple_hash(value):
    # Convert anything to string then bytes
    s = str(value).encode("utf-8")
    # Create SHA-256 hash
    h = hashlib.sha256(s).hexdigest()
    # Keep only alphanumeric chars
    alnum = re.sub(r'[^A-Za-z0-9]', '', h)
    # Return first 19 chars
    return alnum[:19]

print(simple_hash("Tentacles_1"))
print(simple_hash("Tentacles_2"))


# generate a random number
import random
random_number = random.random()
print(random_number)