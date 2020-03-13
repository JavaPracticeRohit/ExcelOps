package errorHandling;

/**
 * Custom Exception class : In case the specific column name is not found in the provided input file
 *
 */
public class ColumnNameNotFound extends Exception{

	public ColumnNameNotFound(String msg){
		super(msg);
	}
}
