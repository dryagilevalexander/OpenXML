﻿using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML
{
    public class ApplicationContext : DbContext
    {
        public DbSet<Condition> Conditions { get; set; }
        public DbSet<SubCondition> SubConditions { get; set; }
        public DbSet<SubConditionParagraph> SubConditionParagraphs { get; set; }
        public DbSet<Contragent> Contragents { get; set; }
        
        public ApplicationContext()
        {
            Database.EnsureDeleted();
            Database.EnsureCreated();
        }

        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            optionsBuilder.UseNpgsql("Host=localhost;Port=5432;Database=openXML;Username=postgres;Password=12345");
        }


        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Contragent>().HasData(
            new Contragent[]
            {
                new Contragent {Id=1, Name = "ООО \"Альфа\"", IsMain = true,  INN ="7701071", KPP="7701001", DirectorName = "Иванов И.И.", DirectorNameR = "Иванова И.И."},
                new Contragent {Id=2, Name = "ООО \"Бетта\"", IsMain = false, INN ="7701051", KPP="7701001", DirectorName = "Капралов Д.М.", DirectorNameR = "Капралова Д.М."},
                new Contragent {Id=3, Name = "ООО \"Гамма\"", IsMain = false, INN ="7701031", KPP="7701002", DirectorName = "Сергеев А.Р.", DirectorNameR = "Сергеева А.Р."},

            });

            modelBuilder.Entity<Condition>().HasData(
            new Condition[]
            {
                new Condition {Id = 1, Name = "Ответственность сторон", TypeOfCondition = 3},
                new Condition {Id = 2, Name = "Предмет договора", TypeOfCondition = 3},
                new Condition {Id = 3, Name = "Права и обязанности сторон", TypeOfCondition = 3}
            });

            modelBuilder.Entity<SubCondition>().HasData(
            new SubCondition[]
            {
                new SubCondition {Id=1, Text="За неисполнение или ненадлежащее исполнение Контракта Стороны несут ответственность в соответствии с законодательством Российской Федерации и условиями Контракта.", ConditionId = 1},
                new SubCondition {Id=2, Text="В случае полного (частичного) неисполнения условий Контракта одной из Сторон эта Сторона обязана возместить другой Стороне причиненные убытки в части, непокрытой неустойкой.", ConditionId = 1},
                new SubCondition {Id=3, Text="В случае просрочки исполнения Подрядчиком обязательств, предусмотренных Контрактом, Подрядчик уплачивает Заказчику пени. Пеня начисляется за каждый день просрочки исполнения Подрядчиком обязательства, предусмотренного Контрактом, начиная со дня, следующего после дня истечения установленного Контрактом срока исполнения обязательства. Размер пени составляет одна трехсотая действующей на дату уплаты пени ключевой ставки Центрального банка Российской Федерации от цены Контракта (отдельного этапа исполнения Контракта), уменьшенной на сумму, пропорциональную объему обязательств, предусмотренных Контрактом (соответствующим отдельным этапом исполнения Контракта) и фактически исполненных Подрядчиком.", ConditionId = 1},
                new SubCondition {Id=4, Text="В случае просрочки исполнения Заказчиком обязательств, предусмотренных Контрактом, Подрядчик вправе потребовать уплату пени в размере одной трехсотой действующей на дату уплаты пеней ключевой ставки Центрального банка Российской Федерации от не уплаченной в срок суммы. Пеня начисляется за каждый день просрочки исполнения обязательства, предусмотренного Контрактом, начиная со дня, следующего после дня истечения установленного Контрактом срока исполнения обязательства.", ConditionId = 1},
                new SubCondition {Id=5, Text="Применение неустойки (штрафа, пени) не освобождает Стороны от исполнения обязательств по Контракту.", ConditionId = 1},
                new SubCondition {Id=6, Text="В случае расторжения Контракта в связи с односторонним отказом Стороны от исполнения Контракта другая Сторона вправе потребовать возмещения только фактически понесенного ущерба, непосредственно обусловленного обстоятельствами, являющимися основанием для принятия решения об одностороннем отказе от исполнения Контракта.", ConditionId = 1},
                new SubCondition {Id=7, Text="Подрядчик обязуется выполнить по заданию Заказчика работу, указанную в пункте 1.2 настоящего договора, и сдать ее результат Заказчику, а Заказчик обязуется принять результат работы и оплатить его.", ConditionId = 2},
                new SubCondition {Id=8, Text="Подрядчик обязуется выполнить следующую работу: <subjectOfContract>, именуемую в дальнейшем \"Работа\".", ConditionId = 2},
                new SubCondition {Id=9, Text="Подрядчик обязуется:", ConditionId = 3},
            });

            modelBuilder.Entity<SubConditionParagraph>().HasData(
            new SubConditionParagraph[]
            {
                new SubConditionParagraph {Id=1, Text="Подрядчик обязуется выполнить Работу с надлежащим качеством, из своих материалов, своими силами и средствами.", SubConditionId = 9},
                new SubConditionParagraph {Id=2, Text="Подрядчик обязуется выполнить Работу в срок до <dateEnd> г.", SubConditionId = 9},

            });
        }
    }
}