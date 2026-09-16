print("=" * 70)
print("BIJOY / SUTONNYMJ CONVERSION TEST")
print("=" * 70)

test_text = "cÖwZ‡e`bmg~n"

print()
print("Original legacy text:")
print(test_text)

print()
print("Character information:")
print()

for i, char in enumerate(test_text, start=1):
    print(
        f"{i:02d} | "
        f"{char!r} | "
        f"Unicode: U+{ord(char):04X}"
    )

print()
print("=" * 70)
print("TEST COMPLETE")
print("=" * 70)