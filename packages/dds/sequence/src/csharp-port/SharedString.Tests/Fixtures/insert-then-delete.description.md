# insert-then-delete

Golden SharedString snapshot fixture for the C# snapshot loader POC.

## Text content

`Hell world!`

## Operations performed

1. Insert `Hello world!` at position 0.
2. Delete positions 4-5 (removes `o`).

## Expected length

11 UTF-16 code units.
