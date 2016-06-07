(Start-Process -FilePath "msiexec.exe" -ArgumentList "/x {716E0306-8318-4364-8B8F-0CC4E9376BAC} /qn" -Wait -Passthru).ExitCode
(Start-Process -FilePath "msiexec.exe" -ArgumentList "/x {A9CF9052-F4A0-475D-A00F-A8388C62DD63} /qn" -Wait -Passthru).ExitCode
(Start-Process -FilePath "msiexec.exe" -ArgumentList "/x {37477865-A3F1-4772-AD43-AAFC6BCFF99F} /qn" -Wait -Passthru).ExitCode
(Start-Process -FilePath "msiexec.exe" -ArgumentList "/x {C04E32E0-0416-434D-AFB9-6969D703A9EF} /qn" -Wait -Passthru).ExitCode
(Start-Process -FilePath "msiexec.exe" -ArgumentList "/x {86493ADD-824D-4B8E-BD72-8C5DCDC52A71} /qn" -Wait -Passthru).ExitCode
(Start-Process -FilePath "msiexec.exe" -ArgumentList "/x {F662A8E6-F4DC-41A2-901E-8C11F044BDEC} /qn" -Wait -Passthru).ExitCode
(Start-Process -FilePath "msiexec.exe" -ArgumentList "/x {196467F1-C11F-4F76-858B-5812ADC83B94} /qn" -Wait -Passthru).ExitCode
(Start-Process -FilePath "msiexec.exe" -ArgumentList "/x {859DFA95-E4A6-48CD-B88E-A3E483E89B44} /qn" -Wait -Passthru).ExitCode
(Start-Process -FilePath "msiexec.exe" -ArgumentList "/x {355B5AC0-CEEE-42C5-AD4D-7F3CFD806C36} /qn" -Wait -Passthru).ExitCode
(Start-Process -FilePath "msiexec.exe" -ArgumentList "/x {1D95BA90-F4F8-47EC-A882-441C99D30C1E} /qn" -Wait -Passthru).ExitCode

Remove-Item -Path "$env:SystemRoot\System32\msxml4*.dll" -Force -ErrorAction SilentlyContinue
Remove-Item -Path "$env:SystemRoot\SysWOW64\msxml4*.dll" -Force -ErrorAction SilentlyContinue