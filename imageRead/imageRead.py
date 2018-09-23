#encoding=utf-8
import base64
from PIL import Image
from PIL import ImageChops


class imageWrap(object):
	def __init__(self,fileName=None):
		self.fileName=fileName
		if self.fileName:
			self.image=self.getImageObject(self.fileName)
		

	def parseImageBase64(self,fileName):
		with open(fileName,'rb') as f:
			f_str=base64.b64encode(f.read())
			return f_str

	def imageSaveAnotherFormat(self,fileName1,fileName2):
		f_str=self.parseImageBase64(fileName1)
		with open(fileName2,"wb") as fp:
			fp.write(base64.b64decode(f_str))
			
	def compareImage(self,fileName1,fileName2):
		f1=self.parseImageBase64(fileName1)
		f2=self.parseImageBase64(fileName2)
		if f1!=f2:
			print "not match"
		else:
			print "match"

	def getImageObject(self,fileName):
		image=Image.open(fileName)
		return image
		
	def returnImageSize(self,imageObj=None):
		if not imageObj:
			return self.image.size
		else:
			return imageObj.size
	def returnImageFormat(self,imageObj=None):
		if not imageObj:
			return self.image.format
		else:
			return imageObj.format
		
	def resizeImage(self,size,imageObj=None):
		if not imageObj:
			return self.image.thumbnail(size)
		else:
			imageObj.thumbnail(size)
		
	def imageSave(self,fileName,imageObj=None):
		if not imageObj:
			return self.image.save(fileName)
		else:
			imageObj.save(fileName)
		
	def comDiff(self,path_one, path_two, diff_save_location):
		"""
		compare two images, if the image are differnt,it will save the different part as another image
		"""
		image_one = Image.open(path_one)
		image_two = Image.open(path_two)
		try: 
			diff = ImageChops.difference(image_one, image_two)
		
			if diff.getbbox() is None:
				print(u"【+】We are the same!")
			else:
				diff.save(diff_save_location)
		except ValueError as e:
			text = (u"The image size may have some issues,pastes another image into this image."
					u"The box argument is either a 2-tuple giving the upper left corner, a 4-tuple defining the left, upper, "
					u"right, and lower pixel coordinate, or None (same as (0, 0)). If a 4-tuple is given, the size of the pasted "
					u"image must match the size of the region.")
			print(u"【{0}】{1}".format(e,text))
	
	def imageCrop(self,x,y,w,h,imageObj=None):
		'''tuple(x presents from the left distance, y presents from the top distance,x+w presents from the left distance+cut width,y+h presents from the up distance+cut height'''
		if not imageObj:
			img_size=self.image.size
			region=self.image.crop((x,y,x+w,y+h))
			return region
		else:
			img_size=imageObj.size
			region=imageObj.crop((x,y,x+w,y+h))
			return region
		
	
	
if __name__=="__main__":
	imageW=imageWrap("1.JPG")
	print imageW.returnImageSize()
	image=imageW.imageCrop(0,0,538,200)
	imageW.imageSave("1.PNG",image)