function Console-Settings {
    ## MAXIMIZE CONSOLE WINDOW ##
    if (-not ([System.Management.Automation.PSTypeName]"Win32.Functions").Type) {
        Add-Type -NameSpace "Win32" -Name "Functions" -MemberDefinition @"
[DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
"@
    }
    [Win32.Functions]::ShowWindow([Win32.Functions]::GetConsoleWindow(), 3) | Out-Null


    ## LOCK CONSOLE SCROLL ##
    $Host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.Size($Host.UI.RawUI.WindowSize.Width, $Host.UI.RawUI.WindowSize.Height)
}


function Write-Centered {
    [CmdletBinding()]
    param(
        # TEXT #
        [Parameter(Position = 0, Mandatory)][string]$Text,
        [ConsoleColor]$TextColor = [ConsoleColor]::White,

        # FRAME #
        [int]$FrameWidth = 60,
        [ConsoleColor]$FrameColor = [ConsoleColor]::DarkRed,

        # LOGO #
        [int]$LogoVerticalMargin = 3,
        [ConsoleColor]$LogoColor = [ConsoleColor]::DarkRed,
        
        # PROMPT #
        [switch]$Prompt_Yes,
        [switch]$Prompt_No,

        # PRODUCT SELECTOR #
        [switch]$ProductSelector
    )
    
    #region ||| HELPERS |||
    
    # MAIN #
    function Pad($len)    { " " * [Math]::Max(0, [int](($Host.UI.RawUI.WindowSize.Width - $len) / 2)) }
    function Center($str) { (Pad $str.Length) + $str }
    function Lines($str)  { $str -split "`n" | % TrimEnd }

    #endregion

    #region ||| SETTINGS |||

    Clear-Host
    Console-Settings

    #endregion

    #region ||| LOGO |||
    Write-Host ("`n" * $LogoVerticalMargin) -NoNewline
    
    $LogoLines = Lines "$LOGO"
    $P = Pad ($LogoLines | Measure-Object -Property Length -Maximum).Maximum
    $LogoLines | % { Write-Host "$P$_" -ForegroundColor "$LogoColor" }

    Write-Host ("`n" * $LogoVerticalMargin) -NoNewline
    
    #endregion

    #region ||| TEXT | FRAME |||
    
    $Frame = Center ("+" + "=" * ($FrameWidth - 2) + "+")
    
    Write-Host ("$Frame`n") -ForegroundColor "$FrameColor"
    Lines "$Text" | % { Write-Host (Center "$_") -ForegroundColor "$TextColor" }
    Write-Host ("`n$Frame") -ForegroundColor "$FrameColor"

    #endregion

    #region ||| PROMPT YES/NO |||
    
    if ($Prompt_Yes -or $Prompt_No) {
        $sel = [int]$Prompt_Yes.IsPresent
        
        Write-Host ("`n" + (Center "SELECT [←/→ | Y/n] | CONFIRM [Enter]") + "`n")
        
        $PromptY = [Console]::CursorTop
        [Console]::CursorVisible = $false

        while ($true) {
            [Console]::SetCursorPosition(0, $PromptY)

            $yColor, $nColor = ("White", "Green")[$sel], ("Red", "White")[$sel]
            $yArrow, $nArrow = (" ", ">")[$sel], (">", " ")[$sel]
            
            $yText = "$yArrow[Y] Yes"
            $nText  = "$nArrow[N] No"
            
            Write-Host (Pad "$yText   $nText".Length) -NoNewline
            Write-Host "$yText" -ForegroundColor "$yColor" -NoNewline
            Write-Host "   $nText" -ForegroundColor "$nColor"
            
            switch (($Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")).VirtualKeyCode) {
                {$_ -in 37,89} { $sel = 1 } # ←, Y
                {$_ -in 39,78} { $sel = 0 } # →, N
                13 { return [bool]$sel }
            }
        }

        [Console]::CursorVisible = $true
    }
    
    #endregion

#region ||| PRODUCT SELECTOR |||

if ($ProductSelector) {
    $products = @(
        @{ Selected = $true ; ID = "Word"       }
        @{ Selected = $true ; ID = "Excel"      }
        @{ Selected = $true ; ID = "PowerPoint" }
        @{ Selected = $false; ID = "Access"     }
    )
    
    $idx = 0
    $itemWidth = 22
    $showWarning = $false

    Write-Host ""
    Write-Host (Center "[↑/↓] NAVIGATE   [← | →] REMOVE | ADD   [SPACE] SWITCH") -ForegroundColor "DarkMagenta"
    Write-Host "`n"
    Write-Host (Center "[A] ALL   [N] NONE   [ENTER] CONFIRM") -ForegroundColor "DarkYellow"
    Write-Host ""

    $ProductSelectorY = [Console]::CursorTop
    [Console]::CursorVisible = $false

    while ($true) {
        [Console]::SetCursorPosition(0, $ProductSelectorY)
        
        for ($i = 0; $i -lt $products.Count; $i++) {
            $box = if ($products[$i].Selected) { "[✓]" } else { "[ ]" }
            $color = if ($products[$i].Selected) { "Green" } else { "White" }
            $name  = $products[$i].ID.PadRight(12)
            
            $padStr = Pad $itemWidth
            
            if ($i -eq $idx) {
                Write-Host $padStr -NoNewline
                Write-Host "► " -NoNewline -ForegroundColor "DarkRed"
                Write-Host "$box $name" -ForegroundColor "$color" -BackgroundColor "DarkGray"
            } else {
                Write-Host "$padStr  $box $name" -ForegroundColor "$color"
            }
        }

        $warnText = if ($showWarning) { Center "⚠  SELECT AT LEAST 1 PRODUCT!" } else { "" }
        Write-Host $warnText.PadRight($Host.UI.RawUI.WindowSize.Width) -ForegroundColor "Red"
        
        switch (($Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")).VirtualKeyCode) {
            38 { $idx = ($idx - 1 + $products.Count) % $products.Count }    # ↑
            40 { $idx = ($idx + 1) % $products.Count }                      # ↓
            37 { $products[$idx].Selected = $false }                        # ←
            39 { $products[$idx].Selected = $true }                         # →
            32 { $products[$idx].Selected = -not $products[$idx].Selected } # SPACE
            65 { $products | ForEach-Object { $_.Selected = $true } }       # A
            78 { $products | ForEach-Object { $_.Selected = $false } }      # N
            13 {
                if ($products | Where-Object { $_.Selected }) {
                    return $products.Where({ -not $_.Selected })
                }
                $showWarning = $true
            }
        }

        if ($products | Where-Object { $_.Selected }) { $showWarning = $false }
    }

    [Console]::CursorVisible = $true
}

#endregion
}


function Get-Image-Path {
    Add-Type -AssemblyName System.Windows.Forms
    $Dialog = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        Title = "Select IMAGE file"
        Filter = "IMAGE files (*.img)|*.img"
        CheckFileExists = $true
    } 
    if ($Dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { 
        return $Dialog.FileName 
    } else { 
        return $null 
    }
}
