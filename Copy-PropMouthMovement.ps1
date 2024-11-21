param(

    $loreditFile = '',
    $xsqFile = '',
    $propName = '',
    $hexMouthColor = 'FFFF0000'

)

$settingsTemplate = "Mix_Average|0|0|full|20|lightorama_colorwash:$($hexMouthColor),1:full,full,single_color|lightorama_none::"

# convert the default row effect into XML
$lorDefaultRowEffectXml = New-Object System.Xml.XmlDocument
$lorDefaultRowEffectXml.LoadXml($lorDefaultRowEffectString)

# Load the S6 sequence
$loreditXml = New-Object xml
$loreditXml.Load($loreditFile)

# Load the xLights sequence
$xsqXml = New-Object xml
$xsqXml.Load($xsqFile)

# Get the voice effects first, there are several here including timing and mouth movements
$voiceEffects = ($xsqXml.xsequence.ElementEffects.Element | Where-Object {$_.name -eq 'voice'}).EffectLayer
$mouthMovementEffects = $null

foreach($voiceEffect in $voiceEffects){

    if($voiceEffect.Effect.label -contains 'etc' -or $voiceEffect -contains 'O' -or $voiceEffect -contains 'AI'){
        $mouthMovementEffects = $voiceEffect.Effect
    }

}

if($null -eq $mouthMovementEffects){
    Write-Error "No mouth movement effects were detected" -ErrorAction Stop
}

# Navigate to the motions rows for the prop that will be using the mouth movements
$prop = $loreditXml.sequence.SequenceProps.SeqProp | Where-Object {$_.name -eq $propName}

if($null -eq $prop){
    Write-Error "No prop matching `'$($propName)`' could be found in the file `'$($loreditFile)`'" -ErrorAction Stop
}

# Get the mouth resting row, this will be needed so we can remove timeslices every time we find another mouth position at the same spot
$propMouthRestingRow = $prop.track | Where-Object {$_.name -eq 'Mouth-rest'}
$propMouthRestingRowOriginalTimeSlice = $propMouthRestingRow.effect
$lastTalkTime = 0

foreach($mouthMovementEffect in $mouthMovementEffects){

    $lorMouthMotionRow = ''
    $thisRowEffectXml = $lorDefaultRowEffectXml.Clone()
    
    switch($mouthMovementEffect.label){

        'AI' {$lorMouthMotionRow = 'Mouth-AI'}
        'E' {$lorMouthMotionRow = 'Mouth-E'}
        'etc' {$lorMouthMotionRow = 'Mouth-etc'}
        'FV' {$lorMouthMotionRow = 'Mouth-FV'}
        'L' {$lorMouthMotionRow = 'Mouth-L'}
        'MBP' {$lorMouthMotionRow = 'Mouth-MBP'}
        'O' {$lorMouthMotionRow = 'Mouth-O'}
        'U' {$lorMouthMotionRow = 'Mouth-U'}
        'WQ' {$lorMouthMotionRow = 'Mouth-WQ'}
        'rest' {$lorMouthMotionRow = 'Mouth-rest'}
        default { Write-Warning "No mouth movement found, will use rest"; $lorMouthMotionRow = 'Mouth-rest'}
    }

    $propMouthMatchingRow = $prop.track | Where-Object {$_.name -eq $lorMouthMotionRow}

    if($null -eq $propMouthMatchingRow){
        Write-Warning "No mouth movement row could be found on the prop. Do you have motion row effects for each mouth movement?"
    }

    # We cant having a resting mouth and talking mouth at the same time, so cut up the resting row
    $newRestEffect = $loreditXml.CreateElement("effect")
    $newRestEffect.SetAttribute("startCentisecond", ($lastTalkTime / 10))
    $newRestEffect.SetAttribute("endCentisecond", ($mouthMovementEffect.startTime / 10))
    $newRestEffect.SetAttribute("intensity", "80")
    $newRestEffect.SetAttribute("settings", $settingsTemplate)

    $propMouthRestingRow.AppendChild($newRestEffect)

    # Create a new effect element
    $newTalkEffect = $loreditXml.CreateElement("effect")
    $newTalkEffect.SetAttribute("startCentisecond", ($mouthMovementEffect.startTime / 10))
    $newTalkEffect.SetAttribute("endCentisecond", ($mouthMovementEffect.endTime / 10))
    $newTalkEffect.SetAttribute("intensity", "80")
    $newTalkEffect.SetAttribute("settings", $settingsTemplate)

    $propMouthMatchingRow.AppendChild($newTalkEffect)

    # Update the last talk time
    $lastTalkTime = $mouthMovementEffect.endTime

}

# Before we write the file back out, get rid of the full mouth resting element
$propMouthRestingRowOriginalTimeSlice.ParentNode.RemoveChild($propMouthRestingRowOriginalTimeSlice)

$loreditXml.Save($loreditFile)
