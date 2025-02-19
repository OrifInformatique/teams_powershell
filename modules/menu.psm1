

# Create a menu that the user can use to select what he want to do
function New-Menu {
    param (
        [System.Collections.Specialized.OrderedDictionary]$dictionary_menu
    )


    [string[]]$move_menu_selector_up_keys = @("W", "UpArrow")
    [string[]]$move_menu_selector_down_keys = @("S", "DownArrow")
    [string[]]$move_menu_selector_select_keys = @("D", "RightArrow", "Enter")
    [string[]]$move_menu_selector_exit_keys = @("Q", "Escape")
    
    [int]$dictionary_size = $dictionary_menu.Count
    [int]$menu_selector = 0
    [bool]$in_menu = $true
    
    while ($in_menu) {

        Clear-Host

        # Display the possible command
        displayAvailableKey
        
        # Display the choices of the menu
        drawMenuChoices -dictionary_menu $dictionary_menu -selector $menu_selector

        # Get the key pressed by the user
        [string]$key_pressed = [console]::ReadKey($true).Key

        switch ($true) {
            ($move_menu_selector_up_keys -contains $key_pressed) {
                $menu_selector--
                break 
            }
            ($move_menu_selector_down_keys -contains $key_pressed) { 
                $menu_selector++
                break 
            }
            ($move_menu_selector_select_keys -contains $key_pressed) { 
                Clear-Host
                # Write-Host " You selected => $($dictionary_menu[$menu_selector])" # DEBUG
                $in_menu = $false
                return $dictionary_menu[$menu_selector]
                break 
            }
            ($move_menu_selector_exit_keys -contains $key_pressed) { 
                Clear-Host
                Write-Host " Exiting Menu..." -ForegroundColor Yellow
                $in_menu = $false
                return 0
                break 
            }
            default { 
                Write-Host " Error in menu" -ForegroundColor Red
                # Write-Host "You pressed: $($key_pressed)"
                break 
            }
        }

        # Put the selector to the top or at the bottom of the menu if he goes out of range of the dictionary
        $menu_selector = keepSelectorInRange -selector $menu_selector -dictionary_size $dictionary_size
    } 
}


# Put the selector to the first or to the last element of the dictionary,
#  if he goes out of range of this dictionary
function keepSelectorInRange {
    [OutputType([int])]
    param (
        [int]$selector,
        [int]$dictionary_size
    )

    if ($selector -gt ($dictionary_size - 1)) {
        $selector = 0
    }
    elseif ($selector -lt 0) {
        $selector = $dictionary_size - 1
    }

    return $selector
}


# Display the dictionary menu keys and the selector
function drawMenuChoices {
    param (
        [System.Collections.Specialized.OrderedDictionary]$dictionary_menu,
        [int]$selector
    )

    [int]$index = 0
    foreach ($key in $dictionary_menu.Keys) {

        if ($index -ne $selector) {
            Write-Host "    $key "
        }
        else {
            Write-Host " -> $key " -BackgroundColor White -ForegroundColor Black
        }
        $index++
    }
}


# TODO: refactor later
# Display the keys the user can use to move trought the menu
function displayAvailableKey {
    param ()
    Write-Host " UP      with [W], [Arrow Up]"
    Write-Host " DOWN    with [S], [Arrow Down]"
    Write-Host " CONFIRM with [D], [Arrow Right], [Enter]"
    Write-Host " QUIT    with [Q], [Escape]"
    Write-Host ""
}

# Visible functions from this module
Export-ModuleMember -Function New-Menu
