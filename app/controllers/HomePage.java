package controllers;

import commons.Tool;
import controllers.base.WebSiteBaseController;
import models.ChangePhoto;
import models.Product;
import models.ProductTypes;
import models.Users;
import play.mvc.With;

import java.util.ArrayList;
import java.util.List;

@With(TopPhotos.class)
public class HomePage extends WebSiteBaseController {

    public static void index() {
        render2();
    }

}