﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace OWL_Site.Models
{
    using System;
    using System.Data.Entity;
    using System.Data.Entity.Infrastructure;
    
    public partial class aspnetdbEntities : DbContext
    {
        public aspnetdbEntities()
            : base("name=aspnetdbEntities")
        {
        }
    
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            throw new UnintentionalCodeFirstException();
        }
    
        public virtual DbSet<AspNetRole> AspNetRoles { get; set; }
        public virtual DbSet<AspNetUserClaim> AspNetUserClaims { get; set; }
        public virtual DbSet<AspNetUserLogin> AspNetUserLogins { get; set; }
        public virtual DbSet<AspNetUserRole> AspNetUserRoles { get; set; }
        public virtual DbSet<AspNetUser> AspNetUsers { get; set; }
        public virtual DbSet<Setting> Settings { get; set; }
        public virtual DbSet<Ivr_Themes> Ivr_Themes { get; set; }
        public virtual DbSet<VmrAlias> VmrAliases { get; set; }
        public virtual DbSet<AllVmr> AllVmrs { get; set; }
        public virtual DbSet<PrivatePhB> PrivatePhBs { get; set; }
        public virtual DbSet<MeetingAttendee> MeetingAttendees { get; set; }
        public virtual DbSet<Meeting> Meetings { get; set; }
    }
}
