import java.util.Date;

/**
 * @author lixizhong
 */
public class Student {

	private String name;
	private String sex;
	private int age;
	private Float scoreYuwen;
    private Float scoreShuxue;
	private Date date;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getSex() {
        return sex;
    }

    public void setSex(String sex) {
        this.sex = sex;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public Float getScoreYuwen() {
        return scoreYuwen;
    }

    public void setScoreYuwen(Float scoreYuwen) {
        this.scoreYuwen = scoreYuwen;
    }

    public Float getScoreShuxue() {
        return scoreShuxue;
    }

    public void setScoreShuxue(Float scoreShuxue) {
        this.scoreShuxue = scoreShuxue;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }
}
