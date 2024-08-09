//This file is automatically generated, please do not modify it manually

using System.Collections.Generic;
using System.Linq;
using NPOI.POIFS.Crypt;

namespace ExcelReader
{
	public partial class Config_Test
	{
		/// <summary>test1</summary>
		public int a = 1;

		/// <summary>test2</summary>
		public long b = 2;

		/// <summary>test3</summary>
		public int c = 3;

		/// <summary>test4</summary>
		public int d = 4;

		/// <summary>test5</summary>
		public string e = "testdesc";

		/// <summary>test6</summary>
		public string f = "simple";

		/// <summary>Test7</summary>
		public float g = 1.2f;

		/// <summary>Test8</summary>
		public bool h = true;

		public static IReadOnlyList<Test> ConfigSheet1 = new List<Test>()
		{
			new Test
			{
				Id = 1,
				t1 = "testt1",
				t2 = new[] { 1, 2, 3, 4, 5, 6 },
				t5 = false,
				t6 = 12.1f,
				t7 = 12.222,
			},
			new Test
			{
				Id = 2,
				t1 = "testt2",
				t2 = new[] { 1, 2, 3, 4, 5, 7 },
				t5 = false,
				t6 = 12.2f,
				t7 = 13.222,
			},
			new Test
			{
				Id = 12,
				t1 = "testt12",
				t2 = new[] { 1, 2, 3, 4, 5, 7 },
				t5 = true,
				t6 = 120.2f,
				t7 = 130.222,
			},
			new Test
			{
				Id = 13,
				t1 = "testt12",
				t2 = new[] { 1, 2, 3, 4, 5, 7 },
				t5 = true,
				t6 = 120.2f,
				t7 = 130.222,
			},
		};
		
		private static readonly Dictionary<int, Dictionary<int, List<Text2>>> _mConfigSheet3 = new();


		private static List<Text2> Asda;

		private static List<Text2> GetAsda()
		{
			if (Asda == null)
			{
				Asda = new List<Text2>()
				{
					new Text2()
					{

					},
				};
			}
			
			return Asda;
		}
		
		public static List<Text2> ConfigSheet3(int key, int key2)
		{
			switch (key)
			{
				case 1:
				{
					if (!_mConfigSheet3.TryGetValue(1, out var ss1))
					{
						ss1 = new Dictionary<int, List<Text2>>();
						_mConfigSheet3[1] = ss1;
					}

					switch (key2)
					{
						case 2:
						{
							if (!ss1.TryGetValue(2, out var ss2))
							{
								ss2 = new List<Text2>()
								{
									new Text2()
									{
										
									},
									new Text2()
									{
										
									},
								};
								ss1[1] = ss2;
							}

							return ss2;
						}
					}
					return null;
				}
			}

			return null;
		}

		private static readonly Dictionary<int, Text2> _mConfigSheet2 = new();
		public static Text2 ConfigSheet2(int key)
		{
			switch (key)
			{
				case 1:
				{
					if (!_mConfigSheet2.ContainsKey(1))
					{
						_mConfigSheet2[1] = new Text2
						{
							t0 = 1,
							t1 = "testt1",
							t2 = new[] { 1, 2, 3, 4, 5, 6 },
							t3 = new[] { "t1", "t2", "t4" },
							t4 = new[] { new[] { "1", "d1" }, new[] { " (2", " d2)" } },
							t5 = false,
							t6 = 12.1f,
							t7 = 12.222,
						};
					}

					return _mConfigSheet2[1];
				}
				case 2:
				{
					if (!_mConfigSheet2.ContainsKey(2))
					{
						_mConfigSheet2[2] = new Text2
						{
							t0 = 2,
							t1 = "testt1",
							t2 = new[] { 1, 2, 3, 4, 5, 6 },
							t3 = new[] { "t1", "t2", "t4" },
							t4 = new[] { new[] { "1", "d1" }, new[] { " (2", " d2)" } },
							t5 = false,
							t6 = 12.1f,
							t7 = 12.222,
						};
					}

					return _mConfigSheet2[2];
				}
			}

			return null;
		}
	}
}
