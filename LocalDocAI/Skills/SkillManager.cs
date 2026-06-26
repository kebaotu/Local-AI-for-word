using LocalDocAI.Models;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace LocalDocAI.Skills
{
    public class SkillManager
    {
        private readonly string _skillsDir;
        private List<Skill> _skills = new List<Skill>();

        public SkillManager(string skillsDirectory)
        {
            _skillsDir = skillsDirectory;
        }

        public void Load()
        {
            _skills.Clear();
            try
            {
                if (!Directory.Exists(_skillsDir)) return;

                foreach (var file in Directory.GetFiles(_skillsDir, "*.json"))
                {
                    try
                    {
                        var json = File.ReadAllText(file);
                        var skill = JsonConvert.DeserializeObject<Skill>(json);
                        if (skill != null && !string.IsNullOrEmpty(skill.Id))
                            _skills.Add(skill);
                    }
                    catch { }
                }
            }
            catch { }
        }

        public List<Skill> GetAll() => _skills.ToList();

        public Skill GetById(string id) => _skills.FirstOrDefault(s => s.Id == id);

        public void Save(Skill skill)
        {
            if (skill == null || string.IsNullOrEmpty(skill.Id)) return;
            try
            {
                Directory.CreateDirectory(_skillsDir);
                var path = Path.Combine(_skillsDir, skill.Id + ".json");
                File.WriteAllText(path, JsonConvert.SerializeObject(skill, Formatting.Indented));
                var existing = _skills.FirstOrDefault(s => s.Id == skill.Id);
                if (existing != null) _skills.Remove(existing);
                _skills.Add(skill);
            }
            catch { }
        }

        public void Delete(string id)
        {
            try
            {
                var path = Path.Combine(_skillsDir, id + ".json");
                if (File.Exists(path)) File.Delete(path);
                _skills.RemoveAll(s => s.Id == id);
            }
            catch { }
        }
    }
}
