# The Northwind Strangler (Fig Pattern)
The purpose off this project is to demonstrate the Strangler Fig pattern on a monolithic Client-Server architecture. I have chosen a legacy application called Northwind, developed by the Microsoft Access team as a sample application to showcase features.

The intention is to apply the strangler pattern to the monolithic application and demonstrate the various patterns involved while also documenting the different phases of implementation, from creating a new client, to removing features and data from the monolith

## What is Northwind?
In 1994 Microsoft invented a sample database model for a fictitious trading company called Northwind, which imports and exports specialty foods from around the world. This model was used to develop templates for their database products to showcase features of the Microsoft toolset. This has led to a variety of example data models and client applications that implement the business requirements, often with monolithic architecture and single source of responsibility data strategy. 

Since its inception Northwind templates have gone through only minor changes and improvements, however just recently a new template called [Northwind 2.0](https://techcommunity.microsoft.com/t5/access-blog/announcing-new-templates-for-microsoft-access-northwind-2-0-is/ba-p/3806082) was released for Microsoft Access Database, with much needed improvements to the data model, UI and documentation.

It is this latest version that is being used in this project to demonstrate the strangler fig pattern.

## What is the Strangler Fig Pattern?
The strangler fig tree, from North East Australia, seeds in the upper branches of a tree and gradually work its way down until it roots in the soil. Over many years the fig tree can grow into fantastic and beautiful shapes, meanwhile strangling and killing the tree that was their host. 

This is used by Martin Fowler as a metaphor to describe a way of re-writing critical systems using the fundamental strategies of Event Interception and Asset Capture. This pattern has been widely adopted for the migration of older monolithic systems as it allows you to break the legacy app into small pieces and deliver the new solution incrementally

For more information on the Strangler pattern in general, Martin Fowler provides a great article, [StranglerFigApplication](https://martinfowler.com/bliki/StranglerFigApplication.html), describing the concept and the patterns involved. For this demonstration, I have adopted an interpretation of the Strangler pattern specifically designed to re-write systems with client-server architecture documented in the article [Patterns - How Would I Strangle Client-Server Applications](https://dowot.gatsbyjs.io/posts/patterns-how-would-i-strangle-client-server-applications)

There are links to other articles detailing the patterns we will use, at the end of this readme

## The Monolith

### Restoring the MS SQL database
#### Prerequisites
MS Sql Server

#### Restoring from backup
A backup of the un-strangled MS SQL database can be found here: `.\ms-access\Northwind 2.0 Dev\Client-Server\Northwind_2.bak`
For the refactored versions you will find a copy in the Northwind Monolith folder of the each phase folder
Follow the microsoft documentation for instruction on [restoring a database from backup](https://learn.microsoft.com/en-us/sql/relational-databases/backup-restore/restore-a-database-backup-using-ssms?view=sql-server-ver15)

TODO

### Running the Northwind monolithic client
#### Prerequisites
MS Office Desktop Suite

#### Opening the Northwind monolithic client with MS Access
The latest copy of the un-strangled version can be found here: `.\ms-access\Northwind 2.0 Dev\Client-Server\Database1.accdb`
For the refactored versions you will find a copy in the Northwind Monolith folder of the each phase folder

### Running the automated tests for Northwind monolithic client
#### Prerequisites
VS Code or Visual Studio
.Net 6 SDK

#### Running tests with the dotnet cli
#### Exploring and Running tests in VS Code
#### Exploring and Running tests with Visual Studio

TODO

## References
TODO