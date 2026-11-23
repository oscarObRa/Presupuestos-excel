#!/usr/bin/env python
"""Test script to run presupuestar.py with simulated input."""
import sys
import os

# Mock stdin
class MockStdin:
    def __init__(self, inputs):
        self.inputs = inputs
        self.index = 0
    
    def readline(self):
        if self.index < len(self.inputs):
            result = self.inputs[self.index] + "\n"
            self.index += 1
            return result
        return "\n"
    
    def __iter__(self):
        return self
    
    def __next__(self):
        if self.index < len(self.inputs):
            result = self.inputs[self.index] + "\n"
            self.index += 1
            return result
        raise StopIteration

# Mock stdin
sys.stdin = MockStdin([
    "CASA VENADO",
    "CLIENTE ICEL",
    "2026-11-21",
    "ICEL-VENADO-001",
    "R01",
    "0.14",
    "0.05",
    "0.0135",
    "0.08",
    "0.16",
    "55",
    "y"
])

# Now run the main
import sys
sys.path.insert(0, ".")

# Import after stdin mock
from presupuestar import main

if __name__ == "__main__":
    # Pass the input file as argument
    main(["H:\\Mi unidad\\ICEL\\PUs\\CASA VENADO\\VENADO.xlsx"])