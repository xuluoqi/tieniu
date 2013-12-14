package controllers.base;

import play.Play;
import play.classloading.enhancers.LocalvariablesNamesEnhancer;
import play.mvc.Controller;
import play.vfs.VirtualFile;

public class WebSiteBaseController extends Controller {

    /**
     * 直接代替之前的render方法，以引入以下规则：
     *   1. 如果在app/views中city{cityId}目录中有模板文件，如上海为app/views/city021/HomePage/index.html，则使用这个模板；如不存在进入下一检查
     *   2. 如果在app/views的defauls有模板文件，如app/views/defaults/HomePage/index.html，则使用这个模板；如不存在进入下一检查
     *   3. 使用app/views/HomePage/index.html；如不存在则抛出异常
     *
     * @param args
     */
    protected static void render2(Object... args) {
        String templateName = null;
        if (args.length > 0 && args[0] instanceof String && LocalvariablesNamesEnhancer.LocalVariablesNamesTracer.getAllLocalVariableNames(args[0]).isEmpty()) {
            templateName = args[0].toString();
        } else {
            templateName = template();
        }

        // 默认模板
        renderTemplate2(templateName, args);
    }

    /**
     * 直接代替之前的renderTemplate方法，以引入以下规则：
     *   1. 如果在app/views中city{cityId}目录中有模板文件，如上海为app/views/city021/HomePage/index.html，则使用这个模板；如不存在进入下一检查
     *   2. 如果在app/views的defauls有模板文件，如app/views/defaults/HomePage/index.html，则使用这个模板；如不存在进入下一检查
     *   3. 使用app/views/HomePage/index.html；如不存在则抛出异常
     *
     * @param templateName
     * @param args
     */
    protected static void renderTemplate2(String templateName, Object... args) {

        // 如果有转入 oldPage=true参数，就直接只render旧的模板

        if (session.get("oldPage") != null || params.get("oldPage") != null) {
            if(Play.mode.isProd()){
                session.put("oldPage", "true");
            }
            renderTemplate(templateName, args);
        }

        String templateRealName = template(templateName);

        // 尝试使用默认模板
        String defaultTemplateName = "defaults/" + templateRealName;
        if (existsTemplateFile(defaultTemplateName)) {
            renderTemplate(defaultTemplateName, args);
        }

        renderTemplate(templateName, args);
    }

    /**
     * 检查模板文件是否存在。
     * TODO: 这一检查应缓存到内存中。
     *
     * @param templateName
     * @return
     */
    private static boolean existsTemplateFile(String templateName) {
        VirtualFile virtualFileTemplate = VirtualFile.fromRelativePath("app/views/" + templateName);
        return virtualFileTemplate.exists();
    }


}
