Attribute VB_Name = "modMath"
Option Explicit

Public Function Ln(n As Double) As Double
    'The ln function returns the natural logarithm of its argument
    'The formula is based on that which is provided by the
    'following web site:
    'http://www.fourmilab.ch/babbage/library.html
    Dim NumberOfTerms As Long
    Dim x As Long
    Dim step1 As Double
    Dim step2 As Double
    
    'The formula is a series calculation and its
    'accuracy improves with a higher number of terms.
    NumberOfTerms = 9999
    step1 = (n - 1) / (n + 1)
    For x = 1 To NumberOfTerms Step 2
        step2 = step2 + ((1 / x) * step1 ^ x)
    Next
    Ln = 2 * step2
End Function

