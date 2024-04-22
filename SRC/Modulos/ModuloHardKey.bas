Attribute VB_Name = "ModuloHardKey"
Declare Function HARDkey Lib "hkey-w32.dll" (ByVal buf As String) As Long

Public sBox1 As Variant
Public sBox2 As Variant
Public password As String

Public Sub InitSB()

password = "bBkkEvEVToeSgkNQ"

sBox1 = Array(&HF8, &HA0, &H32, &HEB, &HC0, &HF2, &HD8, &H16, &H70, &HE, &H22, &HDA, &H44, &H85, &HC2, &H8D, _
&H68, &H74, &H7E, &H3F, &H54, &H31, &HE2, &H5B, &H38, &H7C, &HAA, &HC6, &H11, &H3C, &H34, &H83, _
&H58, &H73, &H1E, &HB8, &H6, &HE1, &H1C, &HA4, &H60, &HF6, &H8E, &HDC, &H4, &HC1, &H30, &HF3, _
&HC8, &HF9, &HFC, &HAF, &H88, &HB1, &H9E, &HB5, &HE8, &HE4, &HB, &H4B, &H26, &HD9, &HD2, &HB4, _
&H23, &H91, &H24, &H25, &HB0, &HA7, &HEC, &HA6, &HA, &H2E, &HD0, &H97, &H1, &HED, &H9C, &H49, _
&H4F, &H12, &H80, &HCA, &H2, &HE3, &H19, &H5D, &HE0, &HE6, &H4C, &H3B, &H50, &H27, &H76, &HAC, _
&H0, &H6E, &HF0, &H6F, &H10, &H4E, &HBA, &H3D, &H5C, &HBB, &H57, &HDF, &H8C, &H98, &H40, &HBF, _
&HA8, &H3, &H9, &HFF, &H48, &H5F, &H7, &H7A, &H61, &HEA, &H94, &H2D, &HC, &HA1, &H62, &H4A, _
&H28, &HBC, &H79, &H8A, &HC4, &H9A, &H5, &H87, &H1D, &HE7, &HF4, &H75, &H6D, &H41, &HD6, &H29, _
&H6C, &H3A, &HD4, &H35, &H2C, &H90, &H6A, &H2F, &H64, &HCF, &H8F, &H45, &H96, &HC7, &HA2, &H77, _
&H42, &HF5, &HCC, &HD5, &H2B, &H71, &H56, &HC9, &H51, &HEF, &H20, &HE5, &H52, &H9D, &H66, &HF7, _
&H18, &H1F, &H59, &H17, &H8, &H93, &H72, &HB6, &H8B, &H2A, &H13, &H95, &H81, &H43, &H78, &HD7, _
&HBE, &H9B, &H92, &HA9, &H5A, &H65, &HDE, &H1A, &HAD, &HCE, &HB2, &H4D, &H69, &HA5, &H36, &H7F, _
&H9F, &H46, &H53, &H33, &H47, &H21, &HD, &HFD, &H7B, &HAB, &HD1, &HF, &H84, &HFA, &HAE, &HCD, _
&H99, &H86, &HBD, &H37, &H14, &H6B, &HDB, &HCB, &HB7, &HFB, &H3E, &H63, &HC5, &HF1, &HFE, &HC3, _
&H82, &H5E, &H55, &H67, &H89, &HA3, &HDD, &H15, &H1B, &HD3, &HB9, &HB3, &H39, &HE9, &HEE, &H7D)

sBox2 = Array(&H78, &HE, &H26, &H1D, &H80, &H77, &HD4, &H1B, &H30, &HE5, &H42, &HC2, &H24, &HD1, &HFA, &HC1, _
&HF8, &HE4, &HDA, &H8A, &H44, &H19, &H36, &H3F, &H98, &H23, &HBC, &HCA, &H18, &H4F, &HDC, &HAC, _
&HB8, &HC8, &H2, &H96, &H28, &H58, &H8C, &HA3, &H88, &H27, &H6C, &H89, &H34, &HDD, &H82, &H13, _
&HE8, &H74, &HE0, &HF4, &H4C, &HD9, &H40, &H46, &HD0, &H93, &H62, &H81, &H14, &H6F, &H5A, &H64, _
&H0, &HA4, &H66, &HE9, &HA8, &H5F, &H86, &H11, &HC, &HB, &H68, &HEE, &H1F, &HE2, &H8E, &H69, _
&H20, &HBF, &H33, &H6B, &HF0, &HBE, &H1, &H3C, &H48, &H79, &H9C, &HD8, &H57, &HFD, &HFE, &HEA, _
&H50, &HEC, &HBA, &H38, &H5C, &H15, &H2F, &H43, &H70, &H9B, &H22, &HA6, &H1A, &H35, &H54, &H8F, _
&H60, &HB5, &HE6, &HC9, &H1C, &HB7, &HAA, &HD5, &H8, &HAE, &H3E, &H6A, &HCC, &HF, &H53, &HF7, _
&H10, &H5E, &HF2, &HD, &H32, &H3A, &HCE, &HCB, &HA0, &H6, &H1E, &HFC, &H16, &H71, &H7B, &H5B, _
&H4, &HE3, &HC4, &H9E, &H52, &H51, &H31, &H75, &H63, &H84, &H21, &HAB, &H90, &H7C, &H6E, &H12, _
&H7A, &HB4, &HB0, &H99, &H94, &H2D, &H56, &H2E, &H39, &HB2, &HF6, &H9D, &H47, &H25, &H37, &H7, _
&H9A, &HA, &H61, &H7D, &H4A, &H8D, &H3D, &HEB, &H17, &H91, &H49, &H29, &H2C, &H92, &HB1, &HDE, _
&HAD, &HB3, &H3, &H7F, &HA2, &HBB, &HC6, &H55, &HA1, &HF5, &HC5, &HFB, &HC7, &HED, &H4D, &H87, _
&HC0, &HA9, &H45, &HAF, &H2A, &H5D, &H76, &H65, &H4E, &HDF, &H41, &H9F, &HD3, &HD7, &H2B, &HC3, _
&HA5, &H5, &H59, &H67, &H83, &HA7, &HD6, &H85, &HB6, &H73, &H3B, &HBD, &H95, &HF3, &H97, &HE1, _
&HE7, &HF1, &H9, &H7E, &H4B, &HB9, &HDB, &HF9, &H72, &HEF, &HCD, &HD2, &H8B, &HFF, &HCF, &H6D)
End Sub
