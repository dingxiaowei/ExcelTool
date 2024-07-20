/*
 * auto generated by tools(注意:千万不要手动修改本文件)
 * avatarguideTest
 */
using System;
using System.IO;
using System.Collections.Generic;
using System.Text;

[Serializable]
public partial class avatarguideTest : IBinarySerializable
{
	/// <summary>
	/// 序号
	/// </summary>
	public int Id { get; set; }
	/// <summary>
	/// gender
	/// </summary>
	public string gender { get; set; }
	/// <summary>
	/// age
	/// </summary>
	public float age { get; set; }
	/// <summary>
	/// bValue
	/// </summary>
	public bool bValue { get; set; }
	/// <summary>
	/// 模型相对中心偏差值
	/// </summary>
	public List<float> vector { get; set; }
	/// <summary>
	/// 格子
	/// </summary>
	public List<List<float>> Grid { get; set; }

	public void DeSerialize(BinaryReader reader)
	{
		Id = reader.ReadInt32();
		gender = reader.ReadString();
		age = reader.ReadSingle();
		bValue = reader.ReadBoolean();
		var vectorCount = reader.ReadInt32();
		if (vectorCount > 0)
		{
			vector = new List<float>();
			for (int i = 0; i < vectorCount; i++)
			{
				vector.Add(reader.ReadSingle());
			}
		}
		else
		{
			vector = null;
		}
		var GridCount = reader.ReadInt32();
		if (GridCount > 0)
		{
			Grid = new List<List<float>>();
			for (int i = 0; i < GridCount; i++)
			{
				var tempList = new List<float>();
				var tempListCount = reader.ReadInt32();
				for (int j = 0; j < tempListCount; j++)
				{
					tempList.Add(reader.ReadSingle());
				}
				Grid.Add(tempList);
			}
		}
		else
		{
			Grid = null;
		}
	}

	public void Serialize(BinaryWriter writer)
	{
		writer.Write(Id);
		writer.Write(gender);
		writer.Write(age);
		writer.Write(bValue);
		if (vector == null || vector.Count == 0)
		{
			writer.Write(0);
		}
		else
		{
			writer.Write(vector.Count);
			for (int i = 0; i < vector.Count; i++)
			{
				writer.Write(vector[i]);
			}
		}
		if (Grid == null || Grid.Count == 0)
		{
			writer.Write(0);
		}
		else
		{
			writer.Write(Grid.Count);
			for (int i = 0; i < Grid.Count; i++)
			{
				writer.Write(Grid[i].Count);
				for (int j = 0; j < Grid[i].Count; j++)
				{
					writer.Write(Grid[i][j]);
				}
			}
		}
	}
}

[Serializable]
public partial class avatarguideTestConfig : IBinarySerializable
{
	Dictionary<int,avatarguideTest> avatarguideTestInfos = new Dictionary<int,avatarguideTest>();
	List<avatarguideTest> avatarguideTestInfoList;

	public List<avatarguideTest> avatarguideTestList()
	{
		if (avatarguideTestInfoList == null)
			avatarguideTestInfoList = new List<avatarguideTest>(avatarguideTestInfos.Values);
		return avatarguideTestInfoList;
	}

	public void DeSerialize(BinaryReader reader)
	{
		int count = reader.ReadInt32();
		for (int i = 0;i < count; i++)
		{
			avatarguideTest tempData = new avatarguideTest();
			tempData.DeSerialize(reader);
			avatarguideTestInfos.Add(tempData.Id, tempData);
		}
	}

	public void Serialize(BinaryWriter writer)
	{
		writer.Write(avatarguideTestInfos.Count);
		for (int i = 0; i < avatarguideTestInfos.Count; i++)
		{
			avatarguideTestInfos[i].Serialize(writer);
		}
	}

	public avatarguideTest QueryById(int id)
	{
		if (avatarguideTestInfos.ContainsKey(id))
			return avatarguideTestInfos[id];
		else
			return null;
	}
}
