package commons;

import java.awt.image.BufferedImage;
import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.imageio.ImageIO;

import com.sun.image.codec.jpeg.JPEGCodec;
import com.sun.image.codec.jpeg.JPEGImageEncoder;

public class ImageUtil {

  	/**
  	 * 改变图片大小
  	 * @param path  // 图片的路径
  	 * @param pathAdd  // 修改后的图片存放路径
  	 * @throws IOException
  	 */
  	public static void changeImgSize(String path,String pathAdd,String fileName,int width,int height) throws IOException{
  	// 将要处理的图片读入   
        BufferedImage img = ImageIO.read(new File(path));   
        File file = new File(pathAdd);
        if(!file.exists())
        	file.mkdirs();
        // 要输出的文件   
        FileOutputStream newImgFile = new FileOutputStream(pathAdd+fileName);   
        // 新建缓存图片对象   
        width = width==0?img.getWidth():width;
        height = height==0?img.getHeight():height;
        BufferedImage newImg = new BufferedImage(width, height,BufferedImage.TYPE_INT_RGB);   
        // 用新建的对象将要处理的图片重新绘制，宽高是原来的一半   
        newImg.getGraphics().drawImage(img, 0, 0, width, height, null);   
           
        JPEGImageEncoder encoder = JPEGCodec.createJPEGEncoder(newImgFile);   
        // 将图片重新编码   
        encoder.encode(newImg);   
        newImgFile.close(); 
  	}
  	
  	/**
  	 * 根据比例改变图片大小
  	 * @param file   用户上传的图片 File file
  	 * @param path   要保存图片的路径  /public/upload/product 
  	 * @param width  要修改的图片宽度 如果 图片宽度大于width 则根据比例压缩 。 否则按照原图片进行上传
  	 * @return
  	 * @throws IOException
  	 */
  	public static String changeImgSize(File file,String path,int width) throws IOException{
  		File toFile = new File(path);
		if (!toFile.exists())
			toFile.mkdirs();
  		String newName = ImageUtil.getNewFileName(file);
  		Integer _width=ImageUtil.getWidthIndexOfRoot(file.getPath());
		Integer _height=ImageUtil.getHeightIndexOfRoot(file.getPath());
		Float f = _width>width?(Float.valueOf(width)/Float.valueOf(_width)):1f;
		changeImgSize(file.getPath(), toFile.getPath()+ "/", newName, (int)(_width*f), (int)(_height*f));
        return 	newName;	
  	}
  	
  	/**
  	 * 根据比例改变图片大小
  	 * @param file   用户上传的图片
  	 * @param path   要保存图片的路径 
  	 * @param width  要修改的图片宽度 如果 图片宽度大于width 则根据比例压缩 。 否则按照原图片进行上传
  	 * @return
  	 * @throws IOException
  	 */
  	public static String changeImgSize(File file,String path,Integer width,Integer height) throws IOException{
  		File toFile = new File(path);
		if (!toFile.exists())
			toFile.mkdirs();
  		String newName = ImageUtil.getNewFileName(file);
		changeImgSize(file.getPath(), toFile.getPath()+ "/", newName, width, height);
        return 	newName;	
  	}
  	
  	/**
  	 
  	 * 根据相对路径获取宽度
  	 * @param path
  	 * @return
  	 */
  	public static Integer getWidth(String path){
  		BufferedImage img=null;
		try {
			img = ImageIO.read(new File(CommonUtil.ROOT_PATH+path));
		} catch (IOException e) {
			e.printStackTrace();
		}   
  		return img.getWidth();
  	}
  	
  	public static Integer getHeight(String path){
  		BufferedImage img=null;
		try {
			img = ImageIO.read(new File(CommonUtil.ROOT_PATH+path));
		} catch (IOException e) {
			e.printStackTrace();
		}   
  		return img.getHeight();
  	}

  	/**
  	 * 根据完整路径获取 图片宽度
  	 * @param path
  	 * @return
  	 */
    public static Integer getWidthIndexOfRoot(String path){
        BufferedImage img=null;
        try {
            img = ImageIO.read(new File(path));
        } catch (IOException e) {
            e.printStackTrace();
        }
        return img.getWidth();
    }

    public static Integer getHeightIndexOfRoot(String path){
        BufferedImage img=null;
        try {
            img = ImageIO.read(new File(path));
        } catch (IOException e) {
            e.printStackTrace();
        }
        return img.getHeight();
    }
  	
  	public static void copy(String oldPath,String toPath){
  			FileInputStream fi=null;
  			BufferedInputStream in=null;
  			FileOutputStream fo=null;
  			BufferedOutputStream out=null;
			try {
			   fi = new FileInputStream(oldPath);
  			   in=new BufferedInputStream(fi);
  			   fo=new FileOutputStream(toPath);
  			   out=new BufferedOutputStream(fo);
  			  
  			  byte[] buf=new byte[1024];
  			  int len=in.read(buf);//读文件，将读到的内容放入到buf数组中，返回的是读到的长度
  			  while(len!=-1){
  			   out.write(buf, 0, len);
  			   len=in.read(buf);
  			  }

  			  out.close();
  			  fo.close();
  			  in.close();
  			  fi.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
  	}
  	
  	 final static void showAllFiles(File dir) throws Exception{
  		  File[] fs = dir.listFiles();
  		  for(int i=0; i<fs.length; i++){
  		   System.out.println(fs[i].getAbsolutePath());
  		   if(fs[i].isDirectory()){
  		    try{
  		     showAllFiles(fs[i]);
  		    }catch(Exception e){}
  		   }
  		  }
  	}
  	
  	/**
  	 * 获取跟上传文件后缀名 一样的新的 时间戳 名字 
  	 * @param file 上传文件
  	 * @return
  	 */
  	public static String getNewFileName(File file){
  		return System.currentTimeMillis()+file.getName().substring(file.getName().indexOf("."));
  	}
}
