using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using Data.Extension;

namespace TestMyLib
{
    [TestFixture]
    public class TestDataExtensions
    {
        private DataTable dtStudent = new DataTable();
        private List<StudentModel> dmStudent = new List<StudentModel>();

        [OneTimeSetUp]
        public void Setup()
        {
            dmStudent = new List<StudentModel>();
            dtStudent = new DataTable();
            dtStudent.Columns.AddRange(
                new DataColumn[3] {
                    new DataColumn("Name"),
                    new DataColumn("StudentId", Type.GetType("System.Int32")),
                    new DataColumn("Age", Type.GetType("System.Int32"))
                }
            );

            dtStudent.Rows.Add("Jack", 15, 100);
            dtStudent.Rows.Add("Smith", 17, 101);
            dtStudent.Rows.Add("Karoro", 20, 102);

            dmStudent.Add(new StudentModel() { Name = "Jack", Age = 15, StudentId = 100 });
            dmStudent.Add(new StudentModel() { Name = "Smith", Age = 17, StudentId = 101 });
            dmStudent.Add(new StudentModel() { Name = "Karoro", Age = 20, StudentId = 102 });
        }

        [Test]
        public void Test01_Table2Model()
        {
            List<StudentModel> dmData = (List<StudentModel>)dtStudent.ToList<StudentModel>();
            if (dmData.Count == 0)
            {
                Assert.Fail("DataModel is empty");
            }
        }

        [Test]
        public void Test02_Table2ModelMapping()
        {
            dtStudent = new DataTable();
            dtStudent.Columns.AddRange(
                new DataColumn[3] {
                    new DataColumn("Name"),
                    new DataColumn("ID", Type.GetType("System.Int32")),
                    new DataColumn("Age", Type.GetType("System.Int32"))
                }
            );

            dtStudent.Rows.Add("Jack", 15, 100);
            dtStudent.Rows.Add("Smith", 17, 101);
            dtStudent.Rows.Add("Karoro", 20, 102);
            List<StudentModel> dmData = (List<StudentModel>)dtStudent.ToList<StudentModel>(new Dictionary<string, string>
            {
                { "Name", "Name" },
                { "StudentId", "ID" },
                { "Age", "Age" },
            });
            if (dmData[0].StudentId == 0)
            {
                Assert.Fail("DataTable Transfer to Class Model with Mapping Fail.");
            }
        }

        [Test]
        public void Test03_Model2Table()
        {
            DataTable dtData = ClassModelConvert.ToDataTable(dmStudent);
            Assert.IsFalse(dtData.Rows.Count == 0);
        }

        [Test]
        public void Test04_Segmentation()
        {
            string test = @"壬戌之秋，七月既望，蘇子與客泛舟遊於赤壁之下。清風徐來，水波不興，舉酒屬客，誦明月之詩，歌窈窕之章。少焉，月出於東山之上，徘徊於斗牛之間，白露橫江，水光接天；縱一葦之所如，凌萬頃之茫然。浩浩乎如馮虛御風，而不知其所止；飄飄乎如遺世獨立，羽化而登仙。

於是飲酒樂甚，扣舷而歌之。歌曰：「桂棹兮蘭槳，擊空明兮泝流光。渺渺兮予懷，望美人兮天一方。」客有吹洞簫者，倚歌而和之，其聲嗚嗚然，如怨如慕，如泣如訴，餘音嫋嫋，不絕如縷。舞幽壑之潛蛟，泣孤舟之嫠婦。

蘇子愀然，正襟危坐，而問客曰：「何為其然也？」

客曰：「『月明星稀，烏鵲南飛』，此非曹孟德之詩乎？西望夏口，東望武昌，山川相繆，鬱乎蒼蒼，此非孟德之困於周郎者乎？方其破荊州，下江陵，順流而東也，舳艫千里，旌旗蔽空，釃酒臨江，橫槊賦詩，固一世之雄也，而今安在哉？況吾與子，漁樵於江渚之上，侶魚蝦而友麋鹿；駕一葉之扁舟，擧匏樽以相屬。寄蜉蝣於天地，渺滄海之一粟。哀吾生之須臾，羨長江之無窮。挾飛仙以遨遊，抱明月而長終。知不可乎驟得，託遺響於悲風。」

蘇子曰：「客亦知夫水與月乎？逝者如斯，而未嘗往也；盈虛者如彼，而卒莫消長也，蓋將自其變者而觀之，則天地曾不能以一瞬；自其不變者而觀之，則物與我皆無盡也，而又何羨乎？且夫天地之間，物各有主，苟非吾之所有，雖一毫而莫取。惟江上之清風，與山間之明月，耳得之而為聲，目遇之而成色，取之無禁，用之不竭，是造物者之無盡藏也，而吾與子之所共適。」

客喜而笑，洗盞更酌。肴核既盡，杯盤狼籍，相與枕藉乎舟中，不知東方之既白。";
            List<TokenModel> _result = ContextIndexing.Segment(test);
        }

        [Test]
        public void Test05_Tokenization()
        {
            string test = @"蘇子與客泛舟遊於赤壁之下";
            List<TokenModel> _result = ContextIndexing.Tokenize(test);
        }
    }
}
