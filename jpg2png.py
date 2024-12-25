from PIL import Image

name = input("Enter the name of the jpg file with a .jpg at the end ")
name2 = input("Enter the name to be saved as ")

img = Image.open(name)  


img.save(name2+".png")
print("Success!")