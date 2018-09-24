
#commande for pushing changes

git commit -m "this is a commitment - %_my_datetime%"

git pull https://TomCou:dtw4571s@github.com/TomCou/sportscontact2.git

#72ce4ac97d95f4c1cd9050061b657c71ae644fc9
git add --all

set _my_datetime=%date%_%time%

git commit -m "this is a commitment - %_my_datetime%"

git push https://TomCou:dtw4571s@github.com/TomCou/sportscontact2.git master --force 

git push heroku master

pause