Imports System.Text

Public Class Files
    ''' <summary>
    ''' Identifies the alteration of the note using the .imy note format expression.
    ''' </summary>
    ''' <param name="expression">IMY note format expression.</param>
    ''' <returns>Return an integer that sums with the note value.</returns>
    ''' <remarks>This value is used to sum with the note and play the corresponding sound.</remarks>
    Public Shared Function GetAlteration(ByVal expression As String) As Note.Alteration
        If expression.Contains("#") = True Then Return Note.Alteration._sustenido
        If expression.Contains("&") = True Then Return Note.Alteration._bemol
        Return Note.Alteration._normal
    End Function
    ''' <summary>
    ''' Identifies the note using the .imy note format expression.
    ''' </summary>
    ''' <param name="expression">IMY note format expression.</param>
    ''' <returns>Return the corresponding note.</returns>
    ''' <remarks>This value can be used to sum with the alteration and play the corresponding sound.</remarks>
    Public Shared Function GetChord(ByVal expression As String) As Note.Notes
        If expression.Contains([Enum].GetName(GetType(Note.Chords), Note.Chords._C).Replace("_", "")) Then Return Note.Notes._dó
        If expression.Contains([Enum].GetName(GetType(Note.Chords), Note.Chords._D).Replace("_", "")) Then Return Note.Notes._ré
        If expression.Contains([Enum].GetName(GetType(Note.Chords), Note.Chords._E).Replace("_", "")) Then Return Note.Notes._mi
        If expression.Contains([Enum].GetName(GetType(Note.Chords), Note.Chords._F).Replace("_", "")) Then Return Note.Notes._fá
        If expression.Contains([Enum].GetName(GetType(Note.Chords), Note.Chords._G).Replace("_", "")) Then Return Note.Notes._sol
        If expression.Contains([Enum].GetName(GetType(Note.Chords), Note.Chords._A).Replace("_", "")) Then Return Note.Notes._lá
        If expression.Contains([Enum].GetName(GetType(Note.Chords), Note.Chords._B).Replace("_", "")) Then Return Note.Notes._si
    End Function
    ''' <summary>
    ''' Interprets which duration the user wants to use when the note is played.
    ''' </summary>
    ''' <param name="expression">IMY note format expression.</param>
    ''' <returns>Integer that represents a note duration.</returns>
    ''' <remarks>The result must be integer and all the types of duration must have a proportion.</remarks>
    Public Shared Function GetDuration(ByVal expression As String) As Integer
        If expression.EndsWith("0.5") Then Return Int((60 / Note.Beat) * 6000)
        If expression.EndsWith("0") Then Return Int((60 / Note.Beat) * 4000)
        If expression.EndsWith("1.5") Then Return Int((60 / Note.Beat) * 3000)
        If expression.EndsWith("1") Then Return Int((60 / Note.Beat) * 2000)
        If expression.EndsWith("2.5") Then Return Int((60 / Note.Beat) * 1500)
        If expression.EndsWith("2") Then Return Int((60 / Note.Beat) * 1000)
        If expression.EndsWith("3.5") Then Return Int((60 / Note.Beat) * 750)
        If expression.EndsWith("3") Then Return Int((60 / Note.Beat) * 500)
        If expression.EndsWith("4.5") Then Return Int((60 / Note.Beat) * 375)
        If expression.EndsWith("4") Then Return Int((60 / Note.Beat) * 250)
        If expression.EndsWith("5.5") Then Return Int((60 / Note.Beat) * 187.5)
        If expression.EndsWith("5") Then Return Int((60 / Note.Beat) * 125)
        '
    End Function
    ''' <summary>
    ''' Reads a .imy file and plays the song.
    ''' </summary>
    ''' <param name="filename">File name path of the song.</param>
    ''' <remarks>This command is called when the program is run from command-line.</remarks>
    Public Shared Sub Read(ByVal filename As String)
        Dim h() As String = IO.File.ReadAllLines(filename, System.Text.Encoding.Default)
        '
        FileOpen(1, filename, OpenMode.Input)
        '
        For Each c In h
            '''' Expressions that may cause an error.
            'If h(0) <> "BEGIN:IMELODY" = True Then
            'MsgBox("O arquivo selecionado não é um arquivo .imy válido.", MsgBoxStyle.Critical, "Arquivo inválido")
            'Exit Sub
            'End If
            'If Trim(c.StartsWith("VERSION:")) = False Then
            'MsgBox("O arquivo selecionado não é um arquivo .imy válido.", MsgBoxStyle.Critical, "Arquivo inválido")
            'Exit Sub
            'End If
            'If Trim(c.StartsWith("FORMAT:CLASS")) = False Then
            'MsgBox("O arquivo selecionado não é um arquivo .imy válido.", MsgBoxStyle.Critical, "Arquivo inválido")
            'Exit Sub
            'End If
            'If Trim(c.StartsWith("MELODY:")) = False Then
            'MsgBox("O arquivo selecionado não é um arquivo .imy válido.", MsgBoxStyle.Critical, "Arquivo inválido")
            'Exit Sub
            'End If
            '
            '''' Expression that makes the music play.
            If Trim(c.StartsWith("BEAT:")) = True Then
                Note.Beat = Int(Split(c, ":")(1))
            End If
            '
            If Trim(c.StartsWith("MELODY:")) = True Then
                '
                Dim notation As String = Split(c, ":")(1).Replace("END", "")
                Dim n As Note
                Dim j As String() = Split(notation, "*")
                '
                Console.ForegroundColor = ConsoleColor.Yellow
                Console.WriteLine("Tocando {0}", filename)
                Console.WriteLine("------------------")
                Console.ForegroundColor = ConsoleColor.Green
                '
                For i As Integer = 1 To UBound(j)
                    n = New Note(GetChord(j(i).ToUpper), GetAlteration(j(i)), Val(j(i)(0)))
                    Console.WriteLine(n.FullName)
                    n.Play()
                    Threading.Thread.Sleep(GetDuration(j(i)))
                Next
                '
                FileClose(1)
                '
                Console.ForegroundColor = ConsoleColor.Yellow
                Console.WriteLine("-------------- FIM DA MÚSICA")
                Console.ForegroundColor = ConsoleColor.White
                '
            End If
            '
        Next
    End Sub

End Class
'
Public Class Note
    ''' <summary>
    ''' Full note name.
    ''' </summary>
    Public Property FullName As String
    ''' <summary>
    ''' Abbreviated note name such as C, D#,F#,G, etc.
    ''' </summary>
    Public Property AbbreviatedName As String
    ''' <summary>
    ''' Note name according .imy format codification.
    ''' </summary>
    Public Property imyName As String
    ''' <summary>
    ''' Integer value of the note.
    ''' </summary>
    Public Property Note As Notes
    ''' <summary>
    ''' .wav file to be opened that corresponds to note name.
    ''' </summary>
    Public Property NoteFile As String
    ''' <summary>
    ''' .wav file converted to byte array.
    ''' </summary>
    Public Property NoteBuffer As Byte() = Text.ASCIIEncoding.Default.GetBytes( _
        "Notes.mkey")
    ''' <summary>
    ''' Octave of the note.
    ''' </summary>
    Public Property Octave As Integer
    ''' <summary>
    ''' Sets or gets the beats per minute (bpm) of the music.
    ''' </summary>
    Public Shared Property Beat As Integer
    ''' <summary>
    ''' Sets or gets the Lenght of the note.
    ''' </summary>
    Public Shared Lenght As Integer
    ''' <summary>
    ''' Represent the musical notes and its relative values.
    ''' </summary>
    ''' <remarks>These values sum with the alteration.</remarks>
    Enum Notes
        _pausa = 0
        _dó = 1
        _ré = 3
        _mi = 5
        _fá = 6
        _sol = 8
        _lá = 10
        _si = 12
    End Enum
    ''' <summary>
    ''' Represent the same musical notes, but with one letter.
    ''' </summary>
    ''' <remarks>These values are used when the 'GetChord' function.</remarks>
    Enum Chords
        _R = 0
        _C = 1
        _D = 3
        _E = 5
        _F = 6
        _G = 8
        _A = 10
        _B = 12
    End Enum
    ''' <summary>
    ''' Represent all the lengths to be used in the variable 'Lenght'.
    ''' </summary>
    ''' <remarks></remarks>
    Enum Lenghts
        semibreve = 0
        mínima = 1
        semínima = 2
        colcheia = 3
        semicolcheia = 4
        fusa = 5
    End Enum
    ''' <summary>
    ''' Gets the alteration signals for sharp and flat notes.
    ''' </summary>
    Public alterationsignals() As Char = {"#", "&"}
    Enum Alteration
        _sustenido = 1
        _normal = 0
        _bemol = -1
    End Enum
    ''' <summary>
    ''' Sets all the properties you need to play a musical note.
    ''' </summary>
    ''' <param name="noteNumber">Note name.</param>
    ''' <param name="alter">Alteration (sharp or flat).</param>
    ''' <param name="octav">Octave (1-8).</param>
    ''' <remarks>This 'Sub' is more useful when the keyword 'New' is included.</remarks>
    Public Sub New(ByVal noteNumber As Notes, ByVal alter As Alteration, ByVal octav As Integer)
        FullName = [Enum].GetName(GetType(Notes), noteNumber).Replace("_", "") & " " & _
            [Enum].GetName(GetType(Alteration), alter).Replace("_", "") & " " & octav

        AbbreviatedName = [Enum].GetName(GetType(Chords), noteNumber).Replace("_", "")
        If alter = Alteration._sustenido Then AbbreviatedName &= alterationsignals(0)
        If alter = Alteration._bemol Then AbbreviatedName &= alterationsignals(1)
        AbbreviatedName &= octav

        imyName = "*" & octav
        If alter = Alteration._sustenido Then imyName &= alterationsignals(0)
        If alter = Alteration._bemol Then imyName &= alterationsignals(1)
        imyName &= [Enum].GetName(GetType(Chords), noteNumber).Replace("_", "").ToLower

        Octave = octav
        Note = noteNumber
        NoteFile = "Notes\note" & noteNumber + alter + (12 * Octave) & ".wav"
    End Sub
    ''' <summary>
    ''' Plays the sound.
    ''' </summary>
    ''' <remarks>The sound are recorded in .wav files, so a message error is 
    ''' displayed if some file is not found.</remarks>
    Public Sub Play()
        If IO.File.Exists(NoteFile) = False Then
            Console.ForegroundColor = ConsoleColor.Red
            Console.WriteLine("Ocorreu um erro ao tentar acessar a nota desejada. " & _
                              "Verifique se o arquivo {0} existe em seu computador.", NoteFile)
            Console.ForegroundColor = ConsoleColor.White
            Exit Sub
        End If
        Dim snd As New Devices.Audio
        snd.Play(NoteFile, AudioPlayMode.Background)
    End Sub

    Public Function GetLength(ByVal duration As Double) As Double
        If duration >= (60 / Beat) * 6000 Then Return 0.5
        If duration >= (60 / Beat) * 4000 And duration < (60 / Beat) * 6000 Then Return 0
        If duration >= (60 / Beat) * 3000 And duration < (60 / Beat) * 4000 Then Return 1.5
        If duration >= (60 / Beat) * 2000 And duration < (60 / Beat) * 3000 Then Return 1
        If duration >= (60 / Beat) * 1500 And duration < (60 / Beat) * 2000 Then Return 2.5
        If duration >= (60 / Beat) * 1000 And duration < (60 / Beat) * 1500 Then Return 2
        If duration >= (60 / Beat) * 750 And duration < (60 / Beat) * 1000 Then Return 3.5
        If duration >= (60 / Beat) * 500 And duration < (60 / Beat) * 750 Then Return 3
        If duration >= (60 / Beat) * 375 And duration < (60 / Beat) * 500 Then Return 4.5
        If duration >= (60 / Beat) * 250 And duration < (60 / Beat) * 375 Then Return 4
        If duration >= (60 / Beat) * 187.5 And duration < (60 / Beat) * 250 Then Return 5.5
        If duration <= (60 / Beat) * 187.5 Then Return 5
    End Function
End Class

Module Mkey
    ''' <summary>
    ''' Determines whether you are saving your music on a .imy file.
    ''' </summary>
    Public Property IsRec As Boolean
    ''' <summary>
    ''' Gets or sets recording file .imy name.
    ''' </summary>
    Public Property RecFile As String

    Sub ShowHeader()
        Console.ForegroundColor = ConsoleColor.Cyan
        Console.WriteLine("{0} {1}", My.Application.Info.Title, My.Application.Info.Version.ToString)
        Console.ForegroundColor = ConsoleColor.Green
        Console.WriteLine("Desenvolvido por {0}", My.Application.Info.CompanyName)
        Console.WriteLine("------------- {0}", My.Application.Info.Copyright)
        Console.WriteLine()
        Console.ForegroundColor = ConsoleColor.White
    End Sub

    Sub Main(ByVal ParamArray args() As String)
        Console.Title = My.Application.Info.Title & " " & My.Application.Info.Version.ToString

        ''''A stopwatch object is created to calculate the note lengths and interpret all the
        ''''    properties to play and save the song.
        Dim T As New Stopwatch
        '
        ShowHeader()
        '
        ''''Try reading the specified file on the command-line, if it really exists.
        Try
            Files.Read(args(1))
        Catch ex As Exception
            ''''Nothing happens if there is an error on reading the file.
        End Try
        '
        ''''Define the default values to start the program and starts the stopwatch.
        Dim n As New Note(Note.Notes._dó, Note.Alteration._normal, 4)
        Note.Beat = 80
        T.Start()
        Console.ForegroundColor = ConsoleColor.White
        '
        ''''The program enters a loop until {ESC} is pressed.
        Do
            Select Case Console.ReadKey.Key
                '''' What is explained under 'Case ConsoleKey.A' is the same
                '''' that occurs on all the key pressing commands (A-J).
                Case ConsoleKey.A
                    '
                    '''' Stopwatch pauses to calculate which note the user wants to play.
                    T.Stop()
                    Note.Lenght = T.ElapsedMilliseconds
                    n.GetLength(Note.Lenght)
                    '
                    '''' Prints the .imy note format expression to the file if {F5} is pressed.
                    If IsRec = True Then Print(1, n.imyName & n.GetLength(Note.Lenght).ToString.Replace(",", "."))
                    '''' Stopwatch restarts to calculate the duration of the next note.
                    T.Restart()
                    '
                    n = New Note(Note.Notes._dó, Note.Alteration._normal, n.Octave)
                    Console.CursorLeft = Console.CursorLeft - 1
                    Console.WriteLine(n.FullName)
                    n.Play()        ' The sound file is played.
                Case ConsoleKey.W
                    T.Stop()
                    Note.Lenght = T.ElapsedMilliseconds
                    n.GetLength(Note.Lenght)
                    If IsRec = True Then Print(1, n.imyName & n.GetLength(Note.Lenght).ToString.Replace(",", "."))
                    T.Restart()
                    '
                    n = New Note(Note.Notes._dó, Note.Alteration._sustenido, n.Octave)
                    Console.CursorLeft = Console.CursorLeft - 1
                    Console.WriteLine(n.FullName)
                    n.Play()
                Case ConsoleKey.S
                    T.Stop()
                    Note.Lenght = T.ElapsedMilliseconds
                    n.GetLength(Note.Lenght)
                    If IsRec = True Then Print(1, n.imyName & n.GetLength(Note.Lenght).ToString.Replace(",", "."))
                    T.Restart()
                    '
                    n = New Note(Note.Notes._ré, Note.Alteration._normal, n.Octave)
                    Console.CursorLeft = Console.CursorLeft - 1
                    Console.WriteLine(n.FullName)
                    n.Play()
                Case ConsoleKey.E
                    T.Stop()
                    Note.Lenght = T.ElapsedMilliseconds
                    n.GetLength(Note.Lenght)
                    If IsRec = True Then Print(1, n.imyName & n.GetLength(Note.Lenght).ToString.Replace(",", "."))
                    T.Restart()
                    '
                    n = New Note(Note.Notes._ré, Note.Alteration._sustenido, n.Octave)
                    Console.CursorLeft = Console.CursorLeft - 1
                    Console.WriteLine(n.FullName)
                    n.Play()
                Case ConsoleKey.D
                    T.Stop()
                    Note.Lenght = T.ElapsedMilliseconds
                    n.GetLength(Note.Lenght)
                    If IsRec = True Then Print(1, n.imyName & n.GetLength(Note.Lenght).ToString.Replace(",", "."))
                    T.Restart()
                    '
                    n = New Note(Note.Notes._mi, Note.Alteration._normal, n.Octave)
                    Console.CursorLeft = Console.CursorLeft - 1
                    Console.WriteLine(n.FullName)
                    n.Play()
                Case ConsoleKey.F
                    T.Stop()
                    Note.Lenght = T.ElapsedMilliseconds
                    n.GetLength(Note.Lenght)
                    If IsRec = True Then Print(1, n.imyName & n.GetLength(Note.Lenght).ToString.Replace(",", "."))
                    T.Restart()
                    '
                    n = New Note(Note.Notes._fá, Note.Alteration._normal, n.Octave)
                    Console.CursorLeft = Console.CursorLeft - 1
                    Console.WriteLine(n.FullName)
                    n.Play()
                Case ConsoleKey.T
                    T.Stop()
                    Note.Lenght = T.ElapsedMilliseconds
                    n.GetLength(Note.Lenght)
                    If IsRec = True Then Print(1, n.imyName & n.GetLength(Note.Lenght).ToString.Replace(",", "."))
                    T.Restart()
                    '
                    n = New Note(Note.Notes._fá, Note.Alteration._sustenido, n.Octave)
                    Console.CursorLeft = Console.CursorLeft - 1
                    Console.WriteLine(n.FullName)
                    n.Play()
                Case ConsoleKey.G
                    T.Stop()
                    Note.Lenght = T.ElapsedMilliseconds
                    n.GetLength(Note.Lenght)
                    If IsRec = True Then Print(1, n.imyName & n.GetLength(Note.Lenght).ToString.Replace(",", "."))
                    T.Restart()
                    '
                    n = New Note(Note.Notes._sol, Note.Alteration._normal, n.Octave)
                    Console.CursorLeft = Console.CursorLeft - 1
                    Console.WriteLine(n.FullName)
                    n.Play()
                Case ConsoleKey.Y
                    T.Stop()
                    Note.Lenght = T.ElapsedMilliseconds
                    n.GetLength(Note.Lenght)
                    If IsRec = True Then Print(1, n.imyName & n.GetLength(Note.Lenght).ToString.Replace(",", "."))
                    T.Restart()
                    '
                    n = New Note(Note.Notes._sol, Note.Alteration._sustenido, n.Octave)
                    Console.CursorLeft = Console.CursorLeft - 1
                    Console.WriteLine(n.FullName)
                    n.Play()
                Case ConsoleKey.H
                    T.Stop()
                    Note.Lenght = T.ElapsedMilliseconds
                    n.GetLength(Note.Lenght)
                    If IsRec = True Then Print(1, n.imyName & n.GetLength(Note.Lenght).ToString.Replace(",", "."))
                    T.Restart()
                    '
                    n = New Note(Note.Notes._lá, Note.Alteration._normal, n.Octave)
                    Console.CursorLeft = Console.CursorLeft - 1
                    Console.WriteLine(n.FullName)
                    n.Play()
                Case ConsoleKey.U
                    T.Stop()
                    Note.Lenght = T.ElapsedMilliseconds
                    n.GetLength(Note.Lenght)
                    If IsRec = True Then Print(1, n.imyName & n.GetLength(Note.Lenght).ToString.Replace(",", "."))
                    T.Restart()
                    '
                    n = New Note(Note.Notes._lá, Note.Alteration._sustenido, n.Octave)
                    Console.CursorLeft = Console.CursorLeft - 1
                    Console.WriteLine(n.FullName)
                    n.Play()
                Case ConsoleKey.J
                    T.Stop()
                    Note.Lenght = T.ElapsedMilliseconds
                    n.GetLength(Note.Lenght)
                    If IsRec = True Then Print(1, n.imyName & n.GetLength(Note.Lenght).ToString.Replace(",", "."))
                    T.Restart()
                    '
                    n = New Note(Note.Notes._si, Note.Alteration._normal, n.Octave)
                    Console.CursorLeft = Console.CursorLeft - 1
                    Console.WriteLine(n.FullName)
                    n.Play()
                    '
                Case ConsoleKey.Spacebar
                    '''' Octave is increased when {SPACEBAR} is pressed.
                    If n.Octave < 8 Then n.Octave += 1
                    Console.WriteLine("Oitava alterada para {0}.", n.Octave)
                    '
                Case ConsoleKey.Backspace
                    '''' Octave is decreased when {BACKSPACE} is pressed.
                    If n.Octave > 1 Then n.Octave -= 1
                    Console.WriteLine("Oitava alterada para {0}.", n.Octave)
                    '
                Case ConsoleKey.Delete
                    '''' All the text is cleared when {DELETE} is pressed.
                    Console.Clear()
                    ShowHeader()
                    If IsRec = True Then Console.ForegroundColor = ConsoleColor.Red
                    '
                Case ConsoleKey.F1
                    '''' Shows the USER MANUAL.
                    Console.ForegroundColor = ConsoleColor.Yellow
                    Console.WriteLine("MANUAL DO USUÁRIO")
                    Console.ForegroundColor = ConsoleColor.White
                    Console.WriteLine()
                    Console.Write("{0} foi desenvolvido para tocar músicas simples com o teclado de seu computador.", My.Application.Info.Title)
                    Console.Write(" Você deve conhecer alguns comandos antes de continuar.")
                    Console.WriteLine(" Continue lendo estas intruções para obter informções precisas sobre o programa.")
                    Console.WriteLine()
                    Console.ForegroundColor = ConsoleColor.Cyan
                    Console.WriteLine("TECLAS e NOTAS ---------------")
                    Console.ForegroundColor = ConsoleColor.White
                    Console.WriteLine()
                    Console.WriteLine("A = dó normal")
                    Console.WriteLine("W = dó sustenido")
                    Console.WriteLine("S = ré normal")
                    Console.WriteLine("E = ré sustenido")
                    Console.WriteLine("D = mi normal")
                    Console.WriteLine("F = fá normal")
                    Console.WriteLine("T = fá sustenido")
                    Console.WriteLine("G = sol normal")
                    Console.WriteLine("Y = sol sustenido")
                    Console.WriteLine("H = lá normal")
                    Console.WriteLine("U = lá sustenido")
                    Console.WriteLine("J = si normal")
                    Console.WriteLine()
                    Console.Write("Esses são os comandos para tocar as sete notas principais.")
                    Console.Write(" No entanto, você precisa mais do que tocar sete notas.")
                    Console.Write(" Um músico usa muitas oitavas numa música.")
                    Console.WriteLine(" Assim, dê atenção às próximas instruções.")
                    Console.WriteLine()
                    Console.WriteLine("BARRA DE ESPAÇO = aumenta 1 oitava")
                    Console.WriteLine("BACKSPACE = diminui 1 oitava")
                    Console.WriteLine()
                    Console.Write("Quando você pressiona a BARRA DE ESPAÇO ou BACKSPACE, uma mensagem é exibida " & _
                                  "para mostrar em qual oitava você está tocando sua música.")
                    Console.Write(" A oitava mais baixa é 1 e a mais alta é 8.")
                    Console.WriteLine(" Sempre que você inicia o programa, o valor da oitava é definido automaticamente para 4, " & _
                                  "porque esta corresponde ao DÓ central do piano.")
                    Console.WriteLine()
                    Console.ForegroundColor = ConsoleColor.Cyan
                    Console.WriteLine("VELOCIDADE DA MÚSICA---------------")
                    Console.ForegroundColor = ConsoleColor.White
                    Console.WriteLine()
                    Console.Write("A velocidade da música é medida em batidas por minuto (bpm).")
                    Console.Write(" A velocidade padrão é 80bpm.")
                    Console.Write(" Você pode mudar esse valor " & _
                                  "apenas quando está iniciando a gravação de sua música.")
                    Console.WriteLine(" É muito importante saber em que velocidade você está tocando, " & _
                                  "pois o programa usará essa velocidade para determinar qual o comprimento correto de suas notas.")
                    Console.WriteLine()
                    Console.ForegroundColor = ConsoleColor.Cyan
                    Console.WriteLine("GRAVANDO SUA MÚSICA ---------------")
                    Console.ForegroundColor = ConsoleColor.White
                    Console.WriteLine()
                    Console.Write("Pressione F5 para iniciar a gravação da música.")
                    Console.WriteLine(" Então você será solicitado a digitar o nome de sua música, " & _
                                  "que será criada no seguinte diretório:")
                    Console.WriteLine()
                    Console.ForegroundColor = ConsoleColor.Yellow
                    Console.WriteLine(My.Computer.FileSystem.SpecialDirectories.MyMusic & "\" & "Mkey")
                    Console.ForegroundColor = ConsoleColor.White
                    Console.WriteLine()
                    Console.Write(" Perceba que o formato usado é "".imy"".")
                    Console.Write(" Esse formato é usado em telefones celulares e é fácil decodificar e salvar.")
                    Console.Write(" Enquanto você grava, as mensagens do programa irão aparecer em vermelho.")
                    Console.WriteLine(" A qualquer momento você pode parar a gravação pressionando F5 novamente.")
                    Console.WriteLine()
                    Console.Write("O programa calcula o tempo que você leva para tocar entre uma nota e outra.")
                    Console.Write(" Baseado neste cálculo e no tempo da música, ele determina " & _
                                  "qual comprimento de nota você deseja usar (ex.: mínima, semínima, colcheia, colcheia pontuada, etc.).")
                    Console.Write(" Assim, se você pressionar duas teclas muito rápido, talvez você consiga uma semifusa.")
                    Console.WriteLine(" Mas se você demorar muito, você conseguirá uma semibreve ou uma mínima.")
                    Console.WriteLine()
                    Console.Write(" Existem limitações neste programa.")
                    Console.Write(" Ele programa reconhece 14 tipos diferentes de comprimento de nota.")
                    Console.Write(" Isso inclui os tipos principais (semibreve, mínima, semínima, colcheia, semicolcheia, " & _
                                  "fusa e semifusa) e suas variações pontuadas.")
                    Console.Write(" Isso significa que o modo como você toca a música pode não ser idêntico à gravação.")
                    Console.WriteLine(" Mas se você seguir corretamente o tempo que você determinou, " & _
                                  "você conseguirá bons resultados.")
                    Console.WriteLine()
                    Console.WriteLine("--------------- FIM DO MANUAL DO USUÁRIO")
                    '
                Case ConsoleKey.F2
                    '''' Shows a list of music saved and asks to choose some music to play.
                    Console.ForegroundColor = ConsoleColor.Green
                    Console.WriteLine("------------------ EXECUTAR MÚSICA DA LISTA")
                    Console.ForegroundColor = ConsoleColor.Yellow
                    '
                    If IO.Directory.Exists(My.Computer.FileSystem.SpecialDirectories.MyMusic & "\Mkey\") = True Then
                        If UBound(IO.Directory.GetFiles(My.Computer.FileSystem.SpecialDirectories.MyMusic & "\Mkey\")) = -1 Then
                            '''' Shows a message if there is no music to delete.
                            Console.WriteLine("Você não tem nenhuma música para excluir.")
                            Console.ForegroundColor = ConsoleColor.White
                            Exit Select
                        End If
                    Else
                        Console.WriteLine("Você não tem nenhuma música para excluir.")
                        Console.ForegroundColor = ConsoleColor.White
                        Exit Select
                    End If
                    '
                    '''' Saves the file names into an array.
                    Dim myfiles As String() = IO.Directory.GetFiles(My.Computer.FileSystem.SpecialDirectories.MyMusic & "\Mkey\")
                    Dim myimy As New Collection()
                    '
                    For c = 0 To UBound(myfiles)
                        '''' The array is loaded to show the file names.
                        Dim i As New IO.FileInfo(myfiles(c))
                        '''' Only the files with .imy extension are shown.
                        If i.Extension = ".imy" Then
                            Console.WriteLine(c + 1 & ". " & i.Name)
                            myimy.Add(i.FullName)
                        End If
                    Next
                    '
                    Console.WriteLine()
                    Console.ForegroundColor = ConsoleColor.Green
                    Console.WriteLine("Digite um nome ou número de arquivo da lista acima:")
                    Console.ForegroundColor = ConsoleColor.White
                    '
                    '''' Starts playing the music.
                    Dim f As String = Console.ReadLine()
                    Try
                        Files.Read(My.Computer.FileSystem.SpecialDirectories.MyMusic & "\Mkey\" & f & ".imy")
                    Catch ex As Exception
                        Try
                            Files.Read(myimy.Item(Int(f)))
                        Catch exe As Exception
                            Console.ForegroundColor = ConsoleColor.Red
                            Console.WriteLine("O arquivo não foi encontrado.")
                            Console.ForegroundColor = ConsoleColor.White
                        End Try
                    End Try
                    '
                Case ConsoleKey.F3
                    '''' Shows a list of musics saved and asks to delete a file.
                    Console.ForegroundColor = ConsoleColor.Red
                    Console.WriteLine("------------------ EXCLUIR MÚSICA DA LISTA")
                    Console.ForegroundColor = ConsoleColor.Yellow
                    '
                    If IO.Directory.Exists(My.Computer.FileSystem.SpecialDirectories.MyMusic & "\Mkey\") = True Then
                        If UBound(IO.Directory.GetFiles(My.Computer.FileSystem.SpecialDirectories.MyMusic & "\Mkey\")) = -1 Then
                            '''' Shows a message if there is no music to delete.
                            Console.WriteLine("Você não tem nenhuma música para excluir.")
                            Console.ForegroundColor = ConsoleColor.White
                            Exit Select
                        End If
                    Else
                        Console.WriteLine("Você não tem nenhuma música para excluir.")
                        Console.ForegroundColor = ConsoleColor.White
                        Exit Select
                    End If
                    '
                    Dim myfiles As String() = IO.Directory.GetFiles(My.Computer.FileSystem.SpecialDirectories.MyMusic & "\Mkey\")
                    Dim myimy As New Collection()
                    '
                    For c = 0 To UBound(myfiles)
                        '''' The array is loaded to show the file names.
                        Dim i As New IO.FileInfo(myfiles(c))
                        '''' Only the files with .imy extension are shown.
                        If i.Extension = ".imy" Then
                            Console.WriteLine(c + 1 & ". " & i.Name)
                            myimy.Add(i.FullName)
                        End If
                    Next
                    '
                    Console.WriteLine()
                    Console.ForegroundColor = ConsoleColor.Green
                    Console.WriteLine("Digite um nome ou número de arquivo da lista acima:")
                    Console.ForegroundColor = ConsoleColor.White
                    '
                    '''' Starts deleting.
                    Dim fx As String = Console.ReadLine()
                    '
                    Console.ForegroundColor = ConsoleColor.Green
                    Console.WriteLine("Você tem certeza que deseja excluir este arquivo? (S/N)")
                    Console.ForegroundColor = ConsoleColor.White
                    '
                    '''' Asks for confirmation to delete the file.
                    If Console.ReadKey.Key = ConsoleKey.S Then
                        '''' Deletes the file if {S} is pressed.
                        Try
                            IO.File.Delete(My.Computer.FileSystem.SpecialDirectories.MyMusic & "\Mkey\" & myimy.Item(Int(fx)) & ".imy")
                        Catch ex As Exception
                            Try
                                IO.File.Delete(myimy.Item(Int(fx)))
                            Catch exe As Exception
                                Console.ForegroundColor = ConsoleColor.Red
                                Console.WriteLine("O arquivo não foi encontrado.")
                                Console.ForegroundColor = ConsoleColor.White
                                Exit Select
                            End Try
                            Console.ForegroundColor = ConsoleColor.Red
                            Console.WriteLine("O arquivo não foi encontrado.")
                            Console.ForegroundColor = ConsoleColor.White
                            Exit Select
                        End Try
                        Console.ForegroundColor = ConsoleColor.Yellow
                        Console.WriteLine("Arquivo excluído!")
                        Console.ForegroundColor = ConsoleColor.White
                        Exit Select
                    End If
                    '
                    '''' The operation is cancelled if {S} is NOT pressed.
                    Console.ForegroundColor = ConsoleColor.Red
                    Console.WriteLine("Operação cancelada.")
                    Console.ForegroundColor = ConsoleColor.White
                    '
                Case ConsoleKey.F5
                        '''' Starts music recording when {F5} is pressed.
                        If IsRec = True Then
                            '''' The music is saved to file only if you've already started recording.
                            IsRec = False
                            RecFile = ""
                            '''' Stopwatch is paused to write the end of the file..
                            T.Stop()
                            '
                            Note.Lenght = T.ElapsedMilliseconds
                            n.GetLength(Note.Lenght)
                            Print(1, n.imyName & n.GetLength(Note.Lenght).ToString.Replace(",", "."))
                            '
                            '''' Stopwatch is reseted to calculate note properties.
                            T.Reset()
                            '
                            Console.WriteLine("--------------")
                            '
                            '''' Ends recording.
                            Print(1, "END:IMELODY")
                            Console.WriteLine("Gravação finalizada.")
                            Console.ForegroundColor = ConsoleColor.White
                            FileClose(1)
                            Console.WriteLine("Arquivo salvo.")
                            Console.WriteLine("---------------")
                            '
                            '''' Stopwatch is restarted.
                            T.Restart()
                            Exit Select
                        End If
                        '
                        '''' Recording loop starts because any music is being recorded.
                        If IsRec = False Then IsRec = True
                        Console.WriteLine("---------------")
                        '''' Set the name of the music.
                        Console.ForegroundColor = ConsoleColor.White
                        Console.WriteLine("Formato de arquivo: .imy")
                        Console.ForegroundColor = ConsoleColor.Green
                        Console.WriteLine("Nome do arquivo: {0}", My.Computer.FileSystem.SpecialDirectories.MyMusic & "\Mkey\")
                        Console.ForegroundColor = ConsoleColor.White
                        RecFile = Console.ReadLine
                        '
                        If RecFile = "" Then
                            '''' Shows an error message because no name has been typed.
                            Console.ForegroundColor = ConsoleColor.Red
                            Console.WriteLine("Você precisa digitar um nome de arquivo!")
                            Console.WriteLine()
                            Console.ForegroundColor = ConsoleColor.White
                            '
                            '''' Recording is cancelled because there is no music file to save.
                            IsRec = False
                            Console.WriteLine("Gravação cancelada.")
                            Exit Select
                        End If
                        '
                        If RecFile.Contains(IO.Path.GetInvalidPathChars) = True Then
                            '''' Shows an error message because invalid characters has been used
                            ''''    in the file name.
                            Console.ForegroundColor = ConsoleColor.Red
                            Console.WriteLine("Você usou caracteres inválidos no caminho. " & _
                                              "Os símbolos {0} não são permitidos.", IO.Path.GetInvalidPathChars)
                            Console.WriteLine()
                            Console.ForegroundColor = ConsoleColor.White
                            '
                            '''' Recording is cancelled because the file name is invalid.
                            IsRec = False
                            Console.WriteLine("Gravação cancelada.")
                            Exit Select
                        End If
                        '
                        '''' Sets the beat per minute of the music. Default: 80bpm.
                        Console.ForegroundColor = ConsoleColor.Green
                        Console.WriteLine("Batidas por minuto (bpm) de sua música: {0}", Note.Beat)
                        Console.ForegroundColor = ConsoleColor.White
                        '
                        '''' The user may type letter and numbers with the beat.
                        ''''    No matter which number is typed, the value is parsed to 80.
                        Try
                            Note.Beat = Int(Console.ReadLine)
                        Catch ex As InvalidCastException
                            Note.Beat = 80
                        End Try
                        '
                        '''' Creates the music directory if it does not exist.
                        If IO.Directory.Exists(My.Computer.FileSystem.SpecialDirectories.MyMusic & "\Mkey\") = False Then
                            IO.Directory.CreateDirectory(My.Computer.FileSystem.SpecialDirectories.MyMusic & "\Mkey\")
                        End If
                        '
                        '''' IMY header is written on the file.
                        FileOpen(1, My.Computer.FileSystem.SpecialDirectories.MyMusic & "\Mkey\" & RecFile & ".imy", 2)
                        PrintLine(1, "BEGIN:IMELODY")
                        PrintLine(1, "VERSION:1.2")
                        PrintLine(1, "BEAT:" & Note.Beat)
                        PrintLine(1, "FORMAT:CLASS1.0")
                        Print(1, "MELODY:")
                        '
                        Console.ForegroundColor = ConsoleColor.Red
                        Console.WriteLine("Gravação iniciada.")
                        Console.WriteLine("---------------")
                        '
                        '''' Stopwatch restarts to record music.
                        T.Restart()
                        '
                Case ConsoleKey.F6
                        ' Header
                        Dim h(13) As Integer
                        h(0) = 77
                        h(1) = 84
                        h(2) = 104
                        h(3) = 100
                        h(4) = 0
                        h(5) = 0
                        h(6) = 0
                        h(7) = 6
                        h(8) = 0
                        h(9) = 1
                        h(10) = 0
                        h(11) = 1
                        h(12) = 0
                        h(13) = 128
                        Dim th(7) As Integer
                        th(0) = 77
                        th(1) = 84
                        th(2) = 114
                        th(3) = 107

                Case ConsoleKey.Escape
                        '''' The program is closed when {ESC} is pressed.
                        Console.ForegroundColor = ConsoleColor.Cyan
                        Console.WriteLine("Finalizando Music Keyboard . . .")
                        '''' Wait for 750 milliseconds to close.
                        Threading.Thread.Sleep(750)
                        Exit Sub
                        '
            End Select
            '
        Loop
        '
    End Sub

End Module
