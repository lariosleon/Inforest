Attribute VB_Name = "modKDS"
Public Sub KDS_AnadirNuevaOrden(ByVal kdsRsCabecera As Recordset)
On Error GoTo ErrorKDSHandler
    Dim Documento As MSXML2.DOMDocument60 ' . DOMDocument60
    Set Documento = New DOMDocument60 'DOMDocument60
    '<Transaction>
    Documento.loadXML ("<Transaction></Transaction>")
    
        '<Order>
        Dim nodoOrder As MSXML2.IXMLDOMNode
        Set nodoOrder = Documento.createElement("Order")
        Documento.documentElement.appendChild nodoOrder
            '<ID>
            Dim nodoID As MSXML2.IXMLDOMNode
            Set nodoID = nodoOrder.appendChild(Documento.createElement("ID"))
            nodoID.Text = Val(Mid(kdsRsCabecera!codigo, 3, Len(kdsRsCabecera!codigo)))
                '<PosTerminal>
                Dim nodoPosTerminal As MSXML2.IXMLDOMNode
                Set nodoPosTerminal = nodoOrder.appendChild(Documento.createElement("PosTerminal"))
                nodoPosTerminal.Text = KDS_Obtener_PosTerminal()
                '<TransType>
                Dim nodoTransType As MSXML2.IXMLDOMNode
                Set nodoTransType = nodoOrder.appendChild(Documento.createElement("TransType"))
                nodoTransType.Text = "1"
                '<OrderStatus>
                Dim nodoOrderStatus As MSXML2.IXMLDOMNode
                Set nodoOrderStatus = nodoOrder.appendChild(Documento.createElement("OrderStatus"))
                nodoOrderStatus.Text = KDS_Obtener_OrderStatus() 'Cambiar
                '<OrderType>
                Dim nodoOrderType As MSXML2.IXMLDOMNode
                Set nodoOrderType = nodoOrder.appendChild(Documento.createElement("OrderType"))
                nodoOrderType.Text = KDS_Obtener_OrderType() 'Cambiar
                '<ServerName>
                Dim nodoServerName As MSXML2.IXMLDOMNode
                Set nodoServerName = nodoOrder.appendChild(Documento.createElement("ServerName"))
                nodoServerName.Text = kdsRsCabecera!tObservacion
                '<Destination>
                Dim nodoDestination As MSXML2.IXMLDOMNode
                Set nodoDestination = nodoOrder.appendChild(Documento.createElement("Destination"))
                nodoDestination.Text = KDS_Obtener_Destination(kdsRsCabecera!tTipoPedido)
                '<GuestTable>
                Dim nodoGuestTable As MSXML2.IXMLDOMNode
                Set nodoGuestTable = nodoOrder.appendChild(Documento.createElement("GuestTable"))
                nodoGuestTable.Text = KDS_Obtener_Mesa(kdsRsCabecera!tMesa)
                '<UserInfo>
                Dim nodoUserInfo As MSXML2.IXMLDOMNode
                Set nodoUserInfo = nodoOrder.appendChild(Documento.createElement("UserInfo"))
                
                Dim kdsRsDetalle As Recordset
                Set kdsRsDetalle = KDS_ObtenerDetallePedido(kdsRsCabecera!codigo)
                                
            Do While Not kdsRsDetalle.EOF
                '<Item>
                Dim nodoItem As MSXML2.IXMLDOMNode
                Set nodoItem = nodoOrder.appendChild(Documento.createElement("Item"))
                    '<ID>
                    Dim nodoItemID As MSXML2.IXMLDOMNode
                    Set nodoItemID = nodoItem.appendChild(Documento.createElement("ID"))
                    If (IsNull(kdsRsDetalle!tItemCombo)) Then
                        nodoItemID.Text = Val(kdsRsDetalle!tItem)
                    Else
                        nodoItemID.Text = Val(kdsRsDetalle!tItem) * 100 + Val(kdsRsDetalle!tItemCombo)
                    End If
                    '<TransType>
                    Dim nodoItemTransType As MSXML2.IXMLDOMNode
                    Set nodoItemTransType = nodoItem.appendChild(Documento.createElement("TransType"))
                    nodoItemTransType.Text = "1"
                    '<Name>
                    Dim nodoItemName As MSXML2.IXMLDOMNode
                    Set nodoItemName = nodoItem.appendChild(Documento.createElement("Name"))
                    
                    If Not kdsRsDetalle!lCombinacion Then
                        nodoItemName.Text = kdsRsDetalle!Producto
                    Else
                        Dim NombreTemp As String
 
                        NombreTemp = KDS_Obtener_InicialesDeNombre(kdsRsDetalle!Combo)
                        nodoItemName.Text = NombreTemp + kdsRsDetalle!Producto
                    End If
                    
                    '<Category>
                    Dim nodoItemCategory As MSXML2.IXMLDOMNode
                    Set nodoItemCategory = nodoItem.appendChild(Documento.createElement("Category"))
                    nodoItemCategory.Text = "Monitor1" ' + kdsRsDetalle!tCodigoProducto
                    '<Quantity>
                    Dim nodoItemQuantity As MSXML2.IXMLDOMNode
                    Set nodoItemQuantity = nodoItem.appendChild(Documento.createElement("Quantity"))
                    nodoItemQuantity.Text = kdsRsDetalle!nCantidad
                    '<Color>
                    Dim nodoItemColor As MSXML2.IXMLDOMNode
                    Set nodoItemColor = nodoItem.appendChild(Documento.createElement("Color"))
                  
                    '<KDSStation>
                    Dim nodoItemKDSStation As MSXML2.IXMLDOMNode
                    Set nodoItemKDSStation = nodoItem.appendChild(Documento.createElement("KDSStation"))
                    nodoItemKDSStation.Text = Val(KDS_ObtenerAreaImpresionKDS(kdsRsDetalle!tCodigoProducto, IIf(IsNull(kdsRsDetalle!tItemCombo), "", kdsRsDetalle!tItemCombo), kdsRsDetalle!tCodigoPedido, kdsRsDetalle!tItem))
                    
                    Dim kdsRsProductoPropiedad As Recordset
                    Set kdsRsProductoPropiedad = KDS_ObtenerPropiedadesProducto(kdsRsDetalle!tCodigoPedido, kdsRsDetalle!tItem, IIf(IsNull(kdsRsDetalle!tItemCombo), "", kdsRsDetalle!tItemCombo), kdsRsDetalle!tCodigoProducto)
                    
                    Do While Not kdsRsProductoPropiedad.EOF
                        '<Condiment>
                        Dim nodoItemCondiment As MSXML2.IXMLDOMNode
                        Set nodoItemCondiment = nodoItem.appendChild(Documento.createElement("Condiment"))
                            '<ID>
                            Dim nodoItemCondimentID As MSXML2.IXMLDOMNode
                            Set nodoItemCondimentID = nodoItemCondiment.appendChild(Documento.createElement("ID"))
                            nodoItemCondimentID.Text = Val(kdsRsProductoPropiedad!tCodigoPropiedad)
                            '<TransType>
                            Dim nodoItemCondimentTransType As MSXML2.IXMLDOMNode
                            Set nodoItemCondimentTransType = nodoItemCondiment.appendChild(Documento.createElement("TransType"))
                            nodoItemCondimentTransType.Text = "1"
                            '<Name>
                            Dim nodoItemCondimentName As MSXML2.IXMLDOMNode
                            Set nodoItemCondimentName = nodoItemCondiment.appendChild(Documento.createElement("Name"))
                            If (kdsRsProductoPropiedad!tCodigoPropiedad <> "9999") Then
                                nodoItemCondimentName.Text = kdsRsProductoPropiedad!Operador + kdsRsProductoPropiedad!Propiedad
                            Else
                                nodoItemCondimentName.Text = KDS_ObtenerInfoPropiedadProducto(kdsRsProductoPropiedad!tEnlace, kdsRsProductoPropiedad!tProducto, kdsRsProductoPropiedad!tCodigoPropiedad)
                            End If
                            '<Color>
                            Dim nodoItemCondimentColor As MSXML2.IXMLDOMNode
                            Set nodoItemCondimentColor = nodoItemCondiment.appendChild(Documento.createElement("Color"))
                            '<Action>
                            Dim nodoItemCondimentAction As MSXML2.IXMLDOMNode
                            Set nodoItemCondimentAction = nodoItemCondiment.appendChild(Documento.createElement("Action"))
                            kdsRsProductoPropiedad.MoveNext
                    Loop
                kdsRsDetalle.MoveNext
            Loop
    Dim direccionArchivo As String
    If Mid(sOrderInfo, Len(sOrderInfo) - 1, 1) <> "\" Then
        sOrderInfo = sOrderInfo + "\"
    End If
    Documento.Save (sOrderInfo & nodoID.Text & ".xml")
    
    Exit Sub
ErrorKDSHandler:
End Sub

Public Sub KDS_EliminarOrden(ByVal kdsRsCabecera As Recordset)
On Error GoTo ErrorKDSHandler
    Dim Documento As MSXML2.DOMDocument60
    Set Documento = New DOMDocument60
    '- <Transaction>
    Documento.loadXML ("<Transaction></Transaction>")
        '- <Order>
        Dim nodoOrder As MSXML2.IXMLDOMNode
        Set nodoOrder = Documento.createElement("Order")
        Documento.documentElement.appendChild nodoOrder
            '<ID>
            Dim nodoID As MSXML2.IXMLDOMNode
            Set nodoID = nodoOrder.appendChild(Documento.createElement("ID"))
            nodoID.Text = Val(Mid(kdsRsCabecera!codigo, 3, Len(kdsRsCabecera!codigo)))
                '<PosTerminal>
                Dim nodoPosTerminal As MSXML2.IXMLDOMNode
                Set nodoPosTerminal = nodoOrder.appendChild(Documento.createElement("PosTerminal"))
                nodoPosTerminal.Text = KDS_Obtener_PosTerminal()
                '<TransType>
                Dim nodoTransType As MSXML2.IXMLDOMNode
                Set nodoTransType = nodoOrder.appendChild(Documento.createElement("TransType"))
                nodoTransType.Text = "2"
                '<ServerName>
                Dim nodoServerName As MSXML2.IXMLDOMNode
                Set nodoServerName = nodoOrder.appendChild(Documento.createElement("ServerName"))
                nodoServerName.Text = kdsRsCabecera!tUsuario
                '<Destination>
                Dim nodoDestination As MSXML2.IXMLDOMNode
                Set nodoDestination = nodoOrder.appendChild(Documento.createElement("Destination"))
                nodoDestination.Text = KDS_Obtener_Destination(kdsRsCabecera!tTipoPedido) 'Cambiar
                '<GuestTable>
                Dim nodoGuestTable As MSXML2.IXMLDOMNode
                Set nodoGuestTable = nodoOrder.appendChild(Documento.createElement("GuestTable"))
                '<UserInfo>
                Dim nodoUserInfo As MSXML2.IXMLDOMNode
                Set nodoUserInfo = nodoOrder.appendChild(Documento.createElement("UserInfo"))
        '</Order>
    '</Transaction>
    Dim direccionArchivo As String
    If Mid(sOrderInfo, Len(sOrderInfo) - 1, 1) <> "\" Then
        sOrderInfo = sOrderInfo + "\"
    End If
    Documento.Save (sOrderInfo & nodoID.Text & ".xml")

    Exit Sub
ErrorKDSHandler:
End Sub

Public Sub KDS_EliminarProducto(ByVal kdsRsCabecera As Recordset, ByVal itItem As String)
On Error GoTo ErrorKDSHandler
    If (KDS_ObtenerProductoPedidoImpresos(kdsRsCabecera!codigo, itItem).RecordCount = 0) Then
        Exit Sub
    End If

    Dim Documento As MSXML2.DOMDocument60
    Set Documento = New DOMDocument60
    '<Transaction>
    Documento.loadXML ("<Transaction></Transaction>")
        '<Order>
        Dim nodoOrder As MSXML2.IXMLDOMNode
        Set nodoOrder = Documento.createElement("Order")
        Documento.documentElement.appendChild nodoOrder
            '<ID>
            Dim nodoID As MSXML2.IXMLDOMNode
            Set nodoID = nodoOrder.appendChild(Documento.createElement("ID"))
            nodoID.Text = Val(Mid(kdsRsCabecera!codigo, 3, Len(kdsRsCabecera!codigo)))
                '<PosTerminal>
                Dim nodoPosTerminal As MSXML2.IXMLDOMNode
                Set nodoPosTerminal = nodoOrder.appendChild(Documento.createElement("PosTerminal"))
                nodoPosTerminal.Text = KDS_Obtener_PosTerminal()
                '<TransType>
                Dim nodoTransType As MSXML2.IXMLDOMNode
                Set nodoTransType = nodoOrder.appendChild(Documento.createElement("TransType"))
                nodoTransType.Text = "3"
                '<OrderStatus>
                Dim nodoOrderStatus As MSXML2.IXMLDOMNode
                Set nodoOrderStatus = nodoOrder.appendChild(Documento.createElement("OrderStatus"))
                nodoOrderStatus.Text = KDS_Obtener_OrderStatus() 'Cambiar
                '<OrderType>
                Dim nodoOrderType As MSXML2.IXMLDOMNode
                Set nodoOrderType = nodoOrder.appendChild(Documento.createElement("OrderType"))
                nodoOrderType.Text = KDS_Obtener_OrderType() 'Cambiar
                '<ServerName>
                Dim nodoServerName As MSXML2.IXMLDOMNode
                Set nodoServerName = nodoOrder.appendChild(Documento.createElement("ServerName"))
                nodoServerName.Text = kdsRsCabecera!tUsuario
                '<Destination>
                Dim nodoDestination As MSXML2.IXMLDOMNode
                Set nodoDestination = nodoOrder.appendChild(Documento.createElement("Destination"))
                nodoDestination.Text = KDS_Obtener_Destination(kdsRsCabecera!tTipoPedido)
                '<GuestTable>
                Dim nodoGuestTable As MSXML2.IXMLDOMNode
                Set nodoGuestTable = nodoOrder.appendChild(Documento.createElement("GuestTable"))
                '<UserInfo>
                Dim nodoUserInfo As MSXML2.IXMLDOMNode
                Set nodoUserInfo = nodoOrder.appendChild(Documento.createElement("UserInfo"))
                
                Dim kdsRsDetalleCombo As Recordset
                Set kdsRsDetalleCombo = KDS_ObtenerDetalleCombo(kdsRsCabecera!codigo, itItem, "0")
            If (kdsRsDetalleCombo.RecordCount = 0) Then 'si no es combo entonces
                '<Item>
                Dim nodoItem As MSXML2.IXMLDOMNode
                Set nodoItem = nodoOrder.appendChild(Documento.createElement("Item"))
                    '<ID>
                    Dim nodoItemID As MSXML2.IXMLDOMNode
                    Set nodoItemID = nodoItem.appendChild(Documento.createElement("ID"))
                    nodoItemID.Text = Val(itItem)
                    '<TransType>
                    Dim nodoItemTransType As MSXML2.IXMLDOMNode
                    Set nodoItemTransType = nodoItem.appendChild(Documento.createElement("TransType"))
                    nodoItemTransType.Text = "2"
            Else 'si no
              Set kdsRsDetalleCombo = KDS_ObtenerDetalleCombo(kdsRsCabecera!codigo, itItem, "1")
              Do While Not kdsRsDetalleCombo.EOF
                '<Item>
                Dim nodoItem2 As MSXML2.IXMLDOMNode
                Set nodoItem2 = nodoOrder.appendChild(Documento.createElement("Item"))
                    '<ID>
                    Dim nodoItemID2 As MSXML2.IXMLDOMNode
                    Set nodoItemID2 = nodoItem2.appendChild(Documento.createElement("ID"))
                    nodoItemID2.Text = Val(kdsRsDetalleCombo!tItem) * 100 + Val(kdsRsDetalleCombo!tItemCombo)
                    '<TransType>
                    Dim nodoItemTransType2 As MSXML2.IXMLDOMNode
                    Set nodoItemTransType2 = nodoItem2.appendChild(Documento.createElement("TransType"))
                    nodoItemTransType2.Text = "2"
                    kdsRsDetalleCombo.MoveNext
              Loop
            End If
    Dim direccionArchivo As String
    If Mid(sOrderInfo, Len(sOrderInfo) - 1, 1) <> "\" Then
        sOrderInfo = sOrderInfo + "\"
    End If
    Documento.Save (sOrderInfo & nodoID.Text & ".xml")
    
    Dim rsProductopedido As Recordset
    Set rsProductopedido = KDS_ObtenerProductoPedido(kdsRsCabecera!codigo, itItem)
    Dim nombreprod As String
    nombreprod = rsProductopedido!tResumido
    Do While Not rsProductopedido.EOF
        If (rsProductopedido!Area <> "") Then
            Call KDS_EnviaMensaje(rsProductopedido!Area, "(" & Val(Mid(kdsRsCabecera!codigo, 3, Len(kdsRsCabecera!codigo))) & ")ELIMINADO:" & nombreprod)
        End If
        rsProductopedido.MoveNext
    Loop
    
    Exit Sub
ErrorKDSHandler:
End Sub

Public Sub KDS_EliminarProductoDeCombo(ByVal kdsRsCabecera As Recordset, ByVal itItem As String, ByVal xItem As String)
On Error GoTo ErrorKDSHandler
    Dim Documento As MSXML2.DOMDocument60
    Set Documento = New DOMDocument60
    '<Transaction>
    Documento.loadXML ("<Transaction></Transaction>")
        '<Order>
        Dim nodoOrder As MSXML2.IXMLDOMNode
        Set nodoOrder = Documento.createElement("Order")
        Documento.documentElement.appendChild nodoOrder
            '<ID>
            Dim nodoID As MSXML2.IXMLDOMNode
            Set nodoID = nodoOrder.appendChild(Documento.createElement("ID"))
            nodoID.Text = Val(Mid(kdsRsCabecera!codigo, 3, Len(kdsRsCabecera!codigo)))
                '<PosTerminal>
                Dim nodoPosTerminal As MSXML2.IXMLDOMNode
                Set nodoPosTerminal = nodoOrder.appendChild(Documento.createElement("PosTerminal"))
                nodoPosTerminal.Text = KDS_Obtener_PosTerminal()
                '<TransType>
                Dim nodoTransType As MSXML2.IXMLDOMNode
                Set nodoTransType = nodoOrder.appendChild(Documento.createElement("TransType"))
                nodoTransType.Text = "3"
                '<OrderStatus>
                Dim nodoOrderStatus As MSXML2.IXMLDOMNode
                Set nodoOrderStatus = nodoOrder.appendChild(Documento.createElement("OrderStatus"))
                nodoOrderStatus.Text = KDS_Obtener_OrderStatus() 'Cambiar
                '<OrderType>
                Dim nodoOrderType As MSXML2.IXMLDOMNode
                Set nodoOrderType = nodoOrder.appendChild(Documento.createElement("OrderType"))
                nodoOrderType.Text = KDS_Obtener_OrderType() 'Cambiar
                '<ServerName>
                Dim nodoServerName As MSXML2.IXMLDOMNode
                Set nodoServerName = nodoOrder.appendChild(Documento.createElement("ServerName"))
                nodoServerName.Text = kdsRsCabecera!tUsuario
                '<Destination>
                Dim nodoDestination As MSXML2.IXMLDOMNode
                Set nodoDestination = nodoOrder.appendChild(Documento.createElement("Destination"))
                nodoDestination.Text = KDS_Obtener_Destination(kdsRsCabecera!tTipoPedido)
                '<GuestTable>
                Dim nodoGuestTable As MSXML2.IXMLDOMNode
                Set nodoGuestTable = nodoOrder.appendChild(Documento.createElement("GuestTable"))
                '<UserInfo>
                Dim nodoUserInfo As MSXML2.IXMLDOMNode
                Set nodoUserInfo = nodoOrder.appendChild(Documento.createElement("UserInfo"))
                
                Dim kdsRsDetalleCombo As Recordset
                Set kdsRsDetalleCombo = KDS_ObtenerDetalleCombo(kdsRsCabecera!codigo, itItem, "0")
                
                '<Item>
                Dim nodoItem2 As MSXML2.IXMLDOMNode
                Set nodoItem2 = nodoOrder.appendChild(Documento.createElement("Item"))
                    '<ID>
                    Dim nodoItemID2 As MSXML2.IXMLDOMNode
                    Set nodoItemID2 = nodoItem2.appendChild(Documento.createElement("ID"))
                    nodoItemID2.Text = Val(itItem) * 100 + Val(xItem)
                    '<TransType>
                    Dim nodoItemTransType2 As MSXML2.IXMLDOMNode
                    Set nodoItemTransType2 = nodoItem2.appendChild(Documento.createElement("TransType"))
                    nodoItemTransType2.Text = "2"
                    kdsRsDetalleCombo.MoveNext

    Dim direccionArchivo As String
    If Mid(sOrderInfo, Len(sOrderInfo) - 1, 1) <> "\" Then
        sOrderInfo = sOrderInfo + "\"
    End If
    Documento.Save (sOrderInfo & nodoID.Text & ".xml")
    
    Dim rsProductopedido As Recordset
    Set rsProductopedido = KDS_ObtenerProductoPedidoDeCombo(kdsRsCabecera!codigo, itItem, xItem)
    Dim nombreprod As String
    nombreprod = rsProductopedido!tResumido
    Dim NombreTemp As String
    NombreTemp = KDS_Obtener_InicialesDeNombre(rsProductopedido!ProtResumido)
    
    If (rsProductopedido!Area <> "") Then
        Call KDS_EnviaMensaje(rsProductopedido!Area, "(" & Val(Mid(kdsRsCabecera!codigo, 3, Len(kdsRsCabecera!codigo))) & ")ELIMINADO:" & NombreTemp & "(" & nombreprod & ")")
    End If
    rsProductopedido.MoveNext
    
    Exit Sub
ErrorKDSHandler:
End Sub

Private Sub KDS_EnviaMensaje(ByVal stationID As String, ByVal info As String)
On Error GoTo ErrorHandler
    Dim Documento As MSXML2.DOMDocument60
    Set Documento = New DOMDocument60
    '<StationInfo>
    Documento.loadXML ("<StationInfo></StationInfo>")
        '<StationID>
        Dim nodoStationID As MSXML2.IXMLDOMNode
        Set nodoStationID = Documento.createElement("StationID")
        nodoStationID.Text = Val(stationID)
        Documento.documentElement.appendChild nodoStationID
        '<User>
        Dim nodoUser As MSXML2.IXMLDOMNode
        Set nodoUser = Documento.createElement("User")
        nodoUser.Text = "0"
        Documento.documentElement.appendChild nodoUser
        '<Info>
        Dim nodoInfo As MSXML2.IXMLDOMNode
        Set nodoInfo = Documento.createElement("Info")
        nodoInfo.Text = info
        Documento.documentElement.appendChild nodoInfo
    
    Dim direccionArchivo As String
    If Mid(sOrderInfo, Len(sOrderInfo) - 1, 1) <> "\" Then
        sOrderInfo = sOrderInfo + "\"
    End If
    Documento.Save (sOrderInfo & "message.xml")
    
    Exit Sub
ErrorHandler:
End Sub

Private Function KDS_Obtener_PosTerminal()
    KDS_Obtener_PosTerminal = Val(sCaja)
End Function

Private Function KDS_Obtener_TransType()
    '<!-- 1 Añadir nuevo orden, para anexar a la última posición-->
    '<!-- 2 Eliminar esta orden. Si utiliza este valor, el KDS sólo hay etiquetas para identificación.
            'Otras etiquetas pueden ser cualquier valor, o no las transferencias -->
    '<!-- 3 Modificar este orden. Sólo tranfer todas las etiquetas para cambiar.
            'Si el valor de la etiqueta está en blanco, KDS tratar con él como sin cambios -->
    '<!-- 4 Reservado para uso futuro -->
    '<!-- 5 Pregunta que este estado de la orden -->
    KDS_Obtener_TransType = "1"
End Function

Private Function KDS_Obtener_OrderStatus()
    '<!-- 0 no pagado -->
    '<!-- 1 pagados -->
    '<!-- 2 En proceso -->
    KDS_Obtener_OrderStatus = "1"
End Function

Private Function KDS_Obtener_OrderType()
    '<!-- ""- Orden normal -->
    '<!-- RUSH- Orden Rush -->
    '<!-- Fire- fire order-->
    KDS_Obtener_OrderType = ""
End Function

Private Function KDS_Obtener_Destination(ByVal tCodigo As String)
    '<!-- El destino de esta orden -->
On Error GoTo ErrorKDSHandler
    Dim destino As String
    destino = Lib.OpenRecordset("USP_KDS_ObtenerTipoPedido '" & tCodigo & "'", Cn)!tDetallado
    KDS_Obtener_Destination = destino
    Exit Function
ErrorKDSHandler:
    Obtener_Destination = ""
End Function

Private Function KDS_Obtener_Category(ByVal tCodigoGrupo As String, ByVal tCodigoSubGrupo As String)
On Error GoTo ErrorKDSHandler
    'Obtener_Category = "A Full Dente Personal"
    Dim categoria As String
    categoria = Lib.OpenRecordset("USP_KDS_ObtenerCategoria '" & tCodigoGrupo & "','" & tCodigoSubGrupo & "'", Cn)!categoria
    KDS_Obtener_Category = categoria
    Exit Function
ErrorKDSHandler:
    KDS_Obtener_Category = ""
End Function

Private Function KDS_ObtenerDetallePedido(ByVal tCodigoPedido As String) As Recordset
On Error GoTo ErrorKDSHandler
    Dim kdsRsDetalle As Recordset
    Set kdsRsDetalle = Lib.OpenRecordset("USP_KDS_ObtenerDetallePedido '" & tCodigoPedido & "'", Cn)
    Set KDS_ObtenerDetallePedido = kdsRsDetalle
    Exit Function
ErrorKDSHandler:
    Set KDS_ObtenerDetallePedido = New Recordset
End Function

Private Function KDS_ObtenerDetalleCombo(ByVal tCodigoPedido As String, ByVal tItem As String, ByVal lImprime As String) As Recordset
On Error GoTo ErrorKDSHandler
    Dim kdsRsDetalleCombo As Recordset
    Set kdsRsDetalleCombo = Lib.OpenRecordset("USP_KDS_ObtenerDetalleCombo '" & tCodigoPedido & "','" & tItem & "','" & lImprime & "'", Cn)
    Set KDS_ObtenerDetalleCombo = kdsRsDetalleCombo
    Exit Function
ErrorKDSHandler:
    Set KDS_ObtenerDetalleCombo = New Recordset
End Function

Private Function KDS_ObtenerPropiedadesProducto(ByVal tCodigoPedido As String, ByVal tItem As String, ByVal tItemCombo As String, ByVal tProducto As String) As Recordset
On Error GoTo ErrorKDSHandler
    Dim kdsRsPropiedadProducto As Recordset
    Set kdsRsPropiedadProducto = Lib.OpenRecordset("USP_KDS_ObtenerPropiedadesProducto '" & tCodigoPedido & "','" & tItem & "','" & tItemCombo & "','" & tProducto & "'", Cn)
    Set KDS_ObtenerPropiedadesProducto = kdsRsPropiedadProducto
    Exit Function
ErrorKDSHandler:
    Set KDS_ObtenerPropiedadesProducto = New Recordset
End Function

Private Function KDS_ObtenerInfoPropiedadProducto(ByVal tEnlace As String, ByVal tCodigoProducto As String, ByVal tCodigoPropiedad As String) As String
On Error GoTo ErrorKDSHandler
    Dim info As Recordset
    Set info = Lib.OpenRecordset("USP_RD_ObtenerInfoPropiedadProducto '" & tEnlace & "','" & tCodigoPropiedad & "','" & tCodigoProducto & "'", Cn)
    info.MoveFirst
    KDS_ObtenerInfoPropiedadProducto = info!Operador + " " + info!Propiedad
    Exit Function
ErrorKDSHandler:
    KDS_ObtenerInfoPropiedadProducto = ""
End Function

Public Function KDS_ValidarProductoArea(ByVal tCodigoProducto As String, ByVal tArea As String) As Boolean
On Error GoTo ErrorKDSHandler
    Dim RsArea As Recordset
    Set RsArea = Lib.OpenRecordset("USP_KDS_ObtenerArea '" & tArea & "'", Cn)
    RsArea.MoveFirst
    
    If (RsArea!KDS = 1) Then
        Dim rsproductoArea As Recordset
        Set rsproductoArea = Lib.OpenRecordset("USP_KDS_ObtenerProductoArea '" & tCodigoProducto & "', '" & tArea & "'", Cn)
        If (rsproductoArea!Cantidad > 0) Then
            KDS_ValidarProductoArea = True 'no se puede insertar
        Else
            KDS_ValidarProductoArea = False 'permiso para insertar
        End If
    Else
        KDS_ValidarProductoArea = False 'permiso para inseratr
    End If
    
    Exit Function
ErrorKDSHandler:
    KDS_ValidarProductoArea = True
End Function

Public Function KDS_ObtenerAreaImpresionKDS(ByVal tCodigoProducto As String, ByVal tItemCombo As String, ByVal tCodigoPedido As String, ByVal tItem As String)
On Error GoTo ErrorKDSHandler
    Dim RsArea As Recordset
    Set RsArea = Lib.OpenRecordset("USP_KDS_ObtenerAreaImpresionKDS '" & tCodigoProducto & "','" & tItemCombo & "','" & tCodigoPedido & "','" & tItem & "'", Cn)
    KDS_ObtenerAreaImpresionKDS = RsArea!tArea
    Exit Function
ErrorKDSHandler:
    KDS_ObtenerAreaImpresionKDS = ""
End Function

Public Function KDS_ObtenerProductoPedido(ByVal tCodigoPedido As String, ByVal tItem As String) As Recordset
On Error GoTo ErrorKDSHandler
    Dim rsProductopedido As Recordset
    Set rsProductopedido = Lib.OpenRecordset("USP_KDS_ObtenerProductoPedido '" & tCodigoPedido & "','" & tItem & "'", Cn)
    Set KDS_ObtenerProductoPedido = rsProductopedido
    Exit Function
ErrorKDSHandler:
    Set KDS_ObtenerProductoPedido = New Recordset
End Function

Public Function KDS_ObtenerProductoPedidoDeCombo(ByVal tCodigoPedido As String, ByVal tItem As String, ByVal xItem As String) As Recordset
On Error GoTo ErrorKDSHandler
    Dim rsProductopedido As Recordset
    Set rsProductopedido = Lib.OpenRecordset("USP_KDS_ObtenerProductoPedidoDeCombo '" & tCodigoPedido & "','" & tItem & "','" & xItem & "'", Cn)
    Set KDS_ObtenerProductoPedidoDeCombo = rsProductopedido
    Exit Function
ErrorKDSHandler:
    Set KDS_ObtenerProductoPedidoDeCombo = New Recordset
End Function

Public Function KDS_ObtenerProductoPedidoImpresos(ByVal tCodigoPedido As String, ByVal tItem As String) As Recordset
On Error GoTo ErrorKDSHandler
    Dim rsProductopedido As Recordset
    Set rsProductopedido = Lib.OpenRecordset("USP_KDS_ObtenerProductoPedidoImpresos '" & tCodigoPedido & "','" & tItem & "'", Cn)
    Set KDS_ObtenerProductoPedidoImpresos = rsProductopedido
    Exit Function
ErrorKDSHandler:
    Set KDS_ObtenerProductoPedidoImpresos = New Recordset
End Function

Private Function KDS_Obtener_Mesa(ByVal tMesa As String) As String
On Error GoTo ErrorKDSHandler
    Dim RsMesa As Recordset
    Set RsMesa = Lib.OpenRecordset("USP_KDS_ObtenerNombreMesaxCodigo '" & tMesa & "'", Cn)
    KDS_Obtener_Mesa = RsMesa!tResumido
    Exit Function
ErrorKDSHandler:
    KDS_Obtener_Mesa = ""
End Function


Public Function KDS_Obtener_InicialesDeNombre(ByVal nombre As String) As String
On Error GoTo ErrorKDSHandler
    Dim NombreTemp As String
    Dim PosNombreTemo As String
    PosNombreTemo = 0
    NombreTemp = Mid(nombre, 1, 1)
    Do
        PosNombreTemo = InStr(PosNombreTemo + 1, nombre, " ", vbTextCompare)
        If (PosNombreTemo) = 0 Then
            Exit Do
        End If
        NombreTemp = NombreTemp & Mid(nombre, PosNombreTemo + 1, 1)
    Loop
    KDS_Obtener_InicialesDeNombre = NombreTemp
    Exit Function
ErrorKDSHandler:
    KDS_Obtener_InicialesDeNombre = ""
End Function

Public Function KDS_ProcesarBumpNotification(ByVal NombArchivo As String)
    Dim strAnio As String
    Dim strMes As String
    Dim strDia As String
    Dim strHora As String
    Dim strMinuto As String
    Dim strSegundo As String
    Dim strPedido As String
    Dim strItemId As String
    
    strAnio = Mid(NombArchivo, 0, 4)
    strMes = Mid(NombArchivo, 5, 2)
    strDia = Mid(NombArchivo, 8, 2)
    strHora = Mid(NombArchivo, 11, 2)
    strMinuto = Mid(NombArchivo, 14, 2)
    strSegundo = Mid(NombArchivo, 0, 4)
    strPedido = Mid(NombArchivo, 0, 4)
    strItemId = Mid(NombArchivo, 0, 4)
End Function
'---------------------------------------------------
    'Agregar lña referencia a Microsoft Scripting Runtime
'---------------------------------------------------
Public Sub KDS_ListarBumpNotification(ByVal sBump As String)
On Error GoTo Err_Sub
    Dim El_Archivo As File
    Dim El_Directorio As Folder
    Set fso = New FileSystemObject
    Set El_Directorio = fso.GetFolder(sBump)
    For Each El_Archivo In El_Directorio.Files
        Dim fFecha As Date
        'El_Archivo.DateCreated
        Cargar_XML sBump & "\" & El_Archivo.Name, El_Archivo.DateCreated
        El_Archivo.Move (sBump & "\Historial\" & El_Archivo.Name)
    Next El_Archivo
Exit Sub
Err_Sub:
MsgBox err.Description
End Sub

Private Sub Cargar_XML(Path_XML As String, FechaCreacion As Date)
On Error GoTo Err_Sub
    Dim objPeopleRoot As IXMLDOMElement
    Dim objPersonElement As IXMLDOMElement
    Dim ObjElement As IXMLDOMNode
    Dim X As IXMLDOMNodeList
    If Len(Dir(Path_XML)) = 0 Then
       MsgBox "El archivo " & Path_XML & _
               " No está en el directorio ." & vbNewLine & _
               " Compruebe la ruta", vbCritical
       Exit Sub
    End If
    'Seteamos la variable
    Set m_objDOMPeople = New DOMDocument60
    m_objDOMPeople.resolveExternals = True
    'Para que valide el documento xml
    m_objDOMPeople.validateOnParse = True
    'Carga el documento
    m_objDOMPeople.async = False
    Call m_objDOMPeople.Load(Path_XML)
    'Comprobamos si se carga
    If m_objDOMPeople.parseError.reason <> "" Then
        ' si hay un error muestra el fallo
        MsgBox m_objDOMPeople.parseError.reason
        Exit Sub
    End If
    Set objPeopleRoot = m_objDOMPeople.documentElement
    Dim Index As Integer
    Dim Lista As IXMLDOMNodeList
    
    Dim ytCodigoPedido As String
    ytCodigoPedido = objPeopleRoot.childNodes.nextNode.childNodes.Item(0).Text
    Dim ytItem As String
    ytItem = objPeopleRoot.childNodes.nextNode.childNodes.Item(8).childNodes.Item(0).Text
    Dim ytCodigoProducto As String
        
    Call USP_KDS_GrabarTiempoSalidaDPedido(ytCodigoPedido, ytItem, ytCodigoProducto, FechaCreacion)
Exit Sub
Err_Sub:
MsgBox err.Description
End Sub

Private Function USP_KDS_GrabarTiempoSalidaDPedido(ByVal tCodigoPedido As String, ByVal tItem As String, ByVal tCodigoProducto As String, ByVal fSalida As Date)
On Error GoTo ErrorKDSHandler
    Dim oComando As clsComando
    Set oComando = New clsComando
    If Not oComando.CreateCmdSp("USP_KDS_GrabarTiempoSalidaDPedido", Cn) Then
        Set oComando = Nothing
        Exit Function
    End If
    oComando.CreateParameter "@tCodigoPedido", adVarChar, adParamInput, 10, tCodigoPedido
    oComando.CreateParameter "@tItem", adVarChar, adParamInput, 3, tItem
    oComando.CreateParameter "@fSalida", adDate, adParamInput, 10, fSalida
    
    If Not oComando.GetParamOK Then
      Set oComando = Nothing
      Exit Function
   End If

   oComando.ExecSP
    Exit Function
ErrorKDSHandler:
    USP_KDS_GrabarTiempoSalidaDPedido = ""
End Function


