rem *** Configure user info ***
git config --global user.name "Dave Smith"
git config --global user.email davsmith@users.noreply.github.com

rem *** Configure tools ***
git config --global core.editor "code --wait"
git config --global diff.tool vscode
git config --global difftool.vscode.cmd "code --wait --diff $LOCAL $REMOTE"
git config --global difftool.prompt false
git config --global merge.tool vscode
git config --global mergetool.vscode.cmd "code --wait $MERGED"
				
rem *** Configure miscellaneous settings ***
git config --global color.branch.upstream "cyan"

rem ***List the settings ***
git config --list --global