Public Function propSig(p1 As Double, p2 As Double, s1 As Integer, s2 As Integer, Optional c As Double = 0.9) As Boolean
	' independent proportion sig-test
	' p1/p2 - proportions 
	' s1/s2 - sample sizes
	' c - confidence level, choices are 80%, 90%(default), 95%
	' If p1 higher or lower than p2, return True, else return False

	Dim score As Double
	Dim threshold As Double

	Let score = (p1 - p2) / Sqr((p1 * (1 - p1)) / s1 + (p2 * (1 - p2)) / s2)

	If c = 0.9 Then
		Let threshold = 1.645
	ElseIf c = 0.8 Then
		Let threshold = 1.28
	ElseIf c = 0.95 Then
		Let threshold = 1.96
	End If

	If Abs(score) >= threshold Then
		pSig = True
		Application.Caller.Font.Color = RGB(237, 125, 49)
	Else
		pSig = False
		Application.Caller.Font.Color = vbBlack
	End If
End Function

Public Function meanSig(m1 As Double, m2 As Double, sd1 As Double, sd2 As Double, n1 As Integer, n2 As Integer, Optional c As Double = 0.9) As Boolean
	' independent t-test
	' m1/m2 - means 
	' sd1/sd2 - standard dev.
	' n1/n2 - sample sizes
	' c - confidence level, choices are 80%, 90%(default), 95%
	' If m1 higher or lower than m2, return True, else return False

	Dim score As Double
	Dim threshold As Double

	Let score = (m1 - m2) / Sqr(sd1 ^ 2 / n1 + sd2 ^ 2 / n2)

	If c = 0.9 Then
		Let threshold = 1.645
	ElseIf c = 0.8 Then
		Let threshold = 1.28
	ElseIf c = 0.95 Then
		Let threshold = 1.96
	End If

	If Abs(score) >= threshold Then
		mSig = True
		Application.Caller.Font.Color = RGB(237, 125, 49)
	Else
		mSig = False
		Application.Caller.Font.Color = vbBlack
	End If
End Function
