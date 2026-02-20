# Setting up GitHub Repository with GitHub CLI

## Quick Setup with `gh` CLI

```bash
# Make sure you're in the project directory
cd /Users/gurudevdutt/CursorProjects/download-gradcas

# Create and push to GitHub in one command
gh repo create download-gradcas --public --source=. --remote=origin --push

# Or if you want it private:
# gh repo create download-gradcas --private --source=. --remote=origin --push
```

## Alternative: Step by step

```bash
# 1. Create the repository (without pushing yet)
gh repo create download-gradcas --public --source=. --remote=origin

# 2. Push your code
git push -u origin main
```

## If repository already exists on GitHub

```bash
# Just add the remote and push
git remote add origin https://github.com/YOUR_USERNAME/download-gradcas.git
git push -u origin main
```

## Check your remotes

```bash
git remote -v
```

## Future workflow

```bash
# Make changes, commit, then push
git add .
git commit -m "Description of changes"
git push
```
