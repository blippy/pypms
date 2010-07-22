rem python -m cProfile -o ..\profile.txt  -s cumulative pypms # >../profile.txt
python -c "import pypms ; pypms.profileit()" >..\profile.txt

