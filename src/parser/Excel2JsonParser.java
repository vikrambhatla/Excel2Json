package parser;

public class Excel2JsonParser {
	public static void main(String [] args)
	{
		GameDataService service = new GameDataService();
		service.searilizeToJson(args[0], args[1]);
	}
}
