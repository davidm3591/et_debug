Hi Sam... I moved the README.md file into a new debug repository. Whenever you push to the remote develop branch, let me know and I will review and merge with master if we are good to go and then update my local debug repo with your changes. Whenever I do updates I will let you know so you can update your local repo ($ git pull). Let me know when you are updating something and I will make sure to let you know when I am updating something, or we could end up with merge conflicts.

https://github.com/davidm3591/et_debug

1. Setup your repository (locally)

2. Navigate to: https://github.com/davidm3591/et_debug

3. Clone repository (I have it in a "projects" folder that has all of my repositories -- c:\projects...)(I am assuming you know how to do this... if you have trouble, let me know)

NOTE: You are going to have setup the develop branch on your local because git/github only clones the master branch by default
   To setup a tracked copy of develop locally:
   $ git fetch
   $ git checkout develop

4. To update content on your local repository:

   Checkout the develop (working) branch (locally)
   $ git checkout develop
   
   Create your new branch locally:
   $ git checkout -b newbranchname offbranchname
   (Example: $ git checkout -b mynewstuffbranch develop)

   If the new branch is not checked out automatically:
   $ git checkout newbranchname

   Complete work then:
   $ git status
   $ git commit -am "Your message"
   OR
   $ git status
   $ git add README.md
   $ git commit -m "Your message"

5. When you are ready to release your changes:
   
   Checkout the develop (working) branch (locally)
   $ git checkout develop
   
   Merge your changes into develop:
   $ git merge newbranchname

6. When you are ready to update remote repository (from your local):

   Update develop (working) branch on github (MAKE SURE YOU ARE ON THE "develop" BRANCH LOCALLY!):
   $ git push

7. When done you can clean up your local repo or leave the new branch for future work:

   To delete your branch from local repo:
   $ git branch -d branchname

NOTE: If you need it you can find a git cheat on github under the same account named: Git - GitBash Cheat Sheet.md
Located: Code-Cheats-Settings/Git-Stuff/Git/Git - GitBash Cheat Sheet.md
  