package controllers;

import models.TopPhoto;
import play.mvc.Before;
import play.mvc.Controller;

/**
 * Created with IntelliJ IDEA.
 * User: upshan
 * Date: 13-10-8
 * Time: 上午11:29
 * To change this template use File | Settings | File Templates.
 */
public class TopPhotos extends Controller {

    @Before
    public static void initTopPhoto() {
        TopPhoto topPhoto = TopPhoto.all().first();
        topPhoto = topPhoto == null ? new TopPhoto() : topPhoto;
        renderArgs.put("topPhoto", topPhoto);
    }
}
