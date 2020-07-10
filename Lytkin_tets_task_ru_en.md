# Взаимодействие с Windows API с помощью PowerShell

Windows API это набор функций, хранящихся в неуправляемых библиотеках DLL операционных систем Microsoft Windows. Интерфейсы Windows API предоставляют доступ к файловой системе, процессам, выводу графической информации и другим функциям ОС. Подробную информацию об интерфейсах Windows API вы найдёте в [документации Microsoft](https://docs.microsoft.com/ru-ru/windows/win32/apiindex/api-index-portal).

Эта статья содержит описание взаимодействия с Windows API с помощью PowerShell-командлета `Add-Type`.

В качестве примера используется функция `CopyFile`, хранящаяся в библиотеке `kernel32.dll`.

>Описанный способ взаимодействия работает только в ОС Windows.

## Calling the `CopyFile` function with the `Add-Type` cmdlet

The `Add-Type` cmdlet allows you to use .NET Framework classes within a PowerShell session.

The cmdlet uses the P/Invoke mechanism to call Windows API functions. The mechanism works with methods with the same C# signatures as the Windows API functions defined in C++. More information on C# signatures of Windows API functions you can find at the [pinvoke.net wiki](http://www.pinvoke.net/).

*To call the `CopyFile` function with the `Add-Type` cmdlet:*

1. Run PowerShell.
2. Import `kernel32.dll` and define the C# signature of the method corresponding to the Windows API function in a new variable:

   ```powershell
   $MethodDefinition = @'
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]

    public static extern bool CopyFile(string lpExistingFileName, string lpNewFileName, bool FailIfExists);
   '@
   ```   

   >The defined method must have public access modifier so you can access it within a PowerShell session.

3. Use the `Add-Type` cmdlet to add the method to a new variable:
   
   ```powershell
   $Kernel32 = Add-Type -MemberDefinition $MethodDefinition -Name "Kernel32" -Namespace "Win32" -PassThru
   ```

   The `PassThru` parameter commands PowerShell to pass the created object.

4. Use the `$Kernel32` object to call the `CopyFile` function:

   ```powershell
   $Kernel32::CopyFile("C:\sample.txt", "C:\example\sample.txt", $False)
   ```

Using the described procedure, you can create functions that implement Windows APIs. See the example:

```powershell
function Copy-RawItem

{
<#
.SYNOPSIS
    Копирует файл из одного каталога в другой.
.PARAMETER Path
    Путь к копируемому файлу.
.PARAMETER Destination
    Целевой путь по которому копируется файл.
.PARAMETER FailIfExists
    Не копировать при обнаружении одноимённого файла в целевом каталоге.
.OUTPUTS
    При успешном копировании отображает объект с информацией о скопированном файле.
.EXAMPLE
    Copy-RawItem -Path C:\sample.txt -Destination C:\example\anyname.txt
#>

    [CmdletBinding()]
    [OutputType([System.IO.FileSystemInfo])]
    Param (
        [Parameter(Mandatory = $True, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Path,
        [Parameter(Mandatory = $True, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [String]
        $Destination,
        [Switch]
        $FailIfExists
    )

    # Импортирование kernel32.dll и создание метода соответвтующего функции CopyFile
    $MethodDefinition = @'
     [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
     public static extern bool CopyFile(string lpExistingFileName, string lpNewFileName, bool bFailIfExists);
'@
   
    $Kernel32 = Add-Type -MemberDefinition $MethodDefinition -Name "Kernel32" -Namespace "Win32" -PassThru

    # Копирование файла с помощью функции CopyFile

    $CopyResult = $Kernel32::CopyFile($Path, $Destination, ([Bool] $PSBoundParameters[‘FailIfExists’]))

    if ($CopyResult -eq $False)
    {
        # Обработка ошибок функции CopyFile. Возвращает исключение ComponentModel.Win32Exception
        throw ( New-Object ComponentModel.Win32Exception )
    }
    else
    {
        Write-Output (Get-ChildItem $Destination)
    }
}
```

## See also:

* [The Add-Type cmdlet's documentation](https://docs.microsoft.com/en-us/powershell/module/Microsoft.PowerShell.Utility/Add-Type?view=powershell-5.1).
