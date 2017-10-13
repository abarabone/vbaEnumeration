VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Enumerable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit







' enumerableSource_ ���ΏۊO�̏ꍇ�́A��̃\�[�X���������� Enumerator ���Ԃ�B

Public Function From(enumerableSource_) As Enumerator
Attribute From.VB_UserMemId = 0
    
    Select Case True
        
        
        Case TypeOf enumerableSource_ Is Enumerator
            
            '�f�ʂ�
            
            Set From = enumerableSource_
            
            
            
        Case TypeOf enumerableSource_ Is Ary
            
            ' Ary �͗񋓎��� EnumVariant ��Ԃ��Ȃ��̂ŁA�O������C�e���[�^���f���Q�[�g�Ƃ��Đݒ肵�Ă��B
            
            Set From = New Enumerator
            
            Dim srcAry_ As Ary:     Set srcAry_ = enumerableSource_
            
            From.SetIterator srcAry_, Delegate.CNew.SetMethod(New EnumOperatorProcs, "OpIteratorAry")
            
            
            
        Case TypeOf enumerableSource_ Is Object , TypeOf enumerableSource_ Is IEnumVARIANT
            
            '���ʂ̃I�u�W�F�N�g�͗񋓎��ɂb�n�l�I�u�W�F�N�g�Ƃ��Ă� EnumVarint ���擾���邽�߁A�I�u�W�F�N�g�����n���΂����B
            
            Set From = New Enumerator
            
            Dim srcUnk_ As IUnknown:  Set srcUnk_ = enumerableSource_
            
            From.SetSource srcUnk_
            
            
            
        Case Else
            
            '�u�����N��Ԃ��B
            
            Set From = New Enumerator
            
            
    End Select
    
End Function











