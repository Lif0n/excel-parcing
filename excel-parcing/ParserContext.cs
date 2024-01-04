using excel_parcing.Models;

using Microsoft.EntityFrameworkCore;

using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace excel_parcing
{
    public class ParserContext : DbContext
    {
        private static ParserContext _instance;

        public static ParserContext Instance
        {
            get
            {
                return _instance ?? (_instance = new ParserContext());
            }
        }

        public DbSet<LessonTeacher> LessonTeachers { get; set; }

        public DbSet<Main_Lesson> MainLessons { get; set; }

        public DbSet<GroupTeacher> GroupTeachers { get; set; }

        public DbSet<TeacherSubject> TeacherSubjects { get; set; }

        public ParserContext()
        {
            Database.EnsureDeleted();
            Database.EnsureCreated();
        }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Main_Lesson>().ToTable("Scheduled-lesson");
        }
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseNpgsql("Server=hnt8.ru;Port=5432;Database=AKVT-Raspisanie;User ID=admin;Password=admin");
        }
    }
}
