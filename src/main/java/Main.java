import com.leticia.ui.Landing;
import javafx.application.Application;
import javafx.scene.Scene;
import javafx.scene.layout.Background;
import javafx.stage.Screen;
import javafx.stage.Stage;

public class Main extends Application {

    @Override
    public void start(Stage stage) throws Exception{

        Landing landing = new Landing(stage);

        // Create the Scene
        Scene scene = new Scene(landing, Screen.getPrimary().getBounds().getWidth() - 200 ,Screen.getPrimary().getBounds().getHeight() - 200 );
        scene.getStylesheets().add( getClass().getResource("application.css").toExternalForm() );

        landing.setBackground( Background.EMPTY );

        // Set the Properties of the Stage
//        stage.setX(0);
//        stage.setY(0);
//        stage.setMinHeight(300);
//        stage.setMinWidth(400);
        stage.centerOnScreen();

        // Add the scene to the Stage
        stage.setScene(scene);
        // Set the title of the Stage
        stage.setTitle("Word2XML converter");
        // Display the Stage
        stage.show();
    }


    public static void main(String[] args) {
        launch(args);
    }
}
